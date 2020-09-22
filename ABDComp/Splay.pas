{*****************************************************************************
  Compress and decompress files using the "splay tree" technique.
  Based on an article by Douglas W. Jones, "Application of Splay Trees to
  Data Compression", in Communications of the ACM, August 1988, page 996.

  This is a method somewhat similar to Huffman encoding (SQZ), but which is
  locally adaptive. It therefore requires only a single pass over the
  uncompressed file, and does not require storage of a code tree with the
  compressed file. It is characterized by code simplicity and low data
  overhead. Compression efficiency is not as good as recent ARC
  implementations, especially for large files. However, for small files, the
  efficiency of SPLAY approaches that of ARC's squashing technique.

  Usage:
    SPLAY [/X] Infile Outfile

    when /X is not specified, Infile is compressed and written to OutFile.
    when /X is specified, InFile must be a file previously compressed by
    SPLAY, and OutFile will contain the expanded text.

    SPLAY will prompt for input if none is given on the command line.

  Caution! This program has very little error checking. It is primarily
  intended as a demonstration of the technique. In particular, SPLAY will
  overwrite OutFile without warning. Speed of SPLAY could be improved
  enormously by writing the inner level bit-processing loops in assembler.

  Implemented on the IBM PC by
    Kim Kokkonen
    TurboPower Software
    [72457,2131]
    8/16/88
*****************************************************************************}

{$R-,S-,I+}

program SplayCompress;

const
  BufSize = 16384;                {Size of input and output buffers}
  Sig = $FF02AA55;                {Arbitrary signature denotes compressed file}

  MaxChar = 256;                  {Ordinal of highest character}
  EofChar = 256;                  {Used to mark end of compressed file}
  PredMax = 255;                  {MaxChar-1}
  TwiceMax = 512;                 {2*MaxChar}
  Root = 0;                       {Index of root node}

  {Used to pack and unpack bits and bytes}
  BitMask : array[0..7] of Byte = (1, 2, 4, 8, 16, 32, 64, 128);

type
  FileHeader =
    record
      Signature : LongInt;
      {Put any other info here, like the original file name or date}
    end;

  BufferArray = array[1..BufSize] of Byte;

  CodeType = 0..MaxChar;          {Has size of Word}
  UpIndex = 0..PredMax;           {Has size of Byte}
  DownIndex = 0..TwiceMax;        {Has size of Word}
  TreeDownArray = array[UpIndex] of DownIndex;
  TreeUpArray = array[DownIndex] of UpIndex;

var
  InBuffer : BufferArray;         {Input file buffer}
  OutBuffer : BufferArray;        {Output file buffer}
  InName : string[79];            {Input file name}
  OutName : string[79];           {Output file name}
  CompStr : string[3];            {Response from Expand? prompt}
  InF : file;                     {Input file}
  OutF : file;                    {Output file}

  Left, Right : TreeDownArray;    {Child branches of code tree}
  Up : TreeUpArray;               {Parent branches of code tree}
  CompressFlag : Boolean;         {True to compress file}
  BitPos : Byte;                  {Current bit in byte}
  InByte : CodeType;              {Current input byte}
  OutByte : CodeType;             {Current output byte}
  InSize : Word;                  {Current chars in input buffer}
  OutSize : Word;                 {Current chars in output buffer}
  Index : Word;                   {General purpose index}

  procedure InitializeSplay;
    {-Initialize the splay tree - as a balanced tree}
  var
    I : DownIndex;
    J : UpIndex;
    K : DownIndex;
  begin
    for I := 1 to TwiceMax do
      Up[I] := (I-1) shr 1;
    for J := 0 to PredMax do begin
      K := (J+1) shl 1;
      Left[J] := K-1;
      Right[J] := K;
    end;
  end;

  procedure Splay(Plain : CodeType);
    {-Rearrange the splay tree for each succeeding character}
  var
    A, B : DownIndex;
    C, D : UpIndex;
  begin
    A := Plain+MaxChar;
    repeat
      {Walk up the tree semi-rotating pairs}
      C := Up[A];
      if C <> Root then begin
        {A pair remains}
        D := Up[C];

        {Exchange children of pair}
        B := Left[D];
        if C = B then begin
          B := Right[D];
          Right[D] := A;
        end else
          Left[D] := A;
        if A = Left[C] then
          Left[C] := B
        else
          Right[C] := B;

        Up[A] := D;
        Up[B] := C;
        A := D;

      end else
        {Handle odd node at end}
        A := C;

    until A = Root;
  end;

  procedure FlushOutBuffer;
    {-Flush output buffer and reset}
  begin
    if OutSize > 0 then begin
      BlockWrite(OutF, OutBuffer, OutSize);
      OutSize := 0;
    end;
  end;

  procedure WriteByte;
    {-Output byte in OutByte}
  begin
    if OutSize = BufSize then
      FlushOutBuffer;
    Inc(OutSize);
    OutBuffer[OutSize] := OutByte;
  end;

  procedure Compress(Plain : CodeType);
    {-Compress a single character}
  var
    A : DownIndex;
    U : UpIndex;
    Sp : 0..MaxChar;
    Stack : array[UpIndex] of Boolean;
  begin
    A := Plain+MaxChar;
    Sp := 0;

    {Walk up the tree pushing bits onto stack}
    repeat
      U := Up[A];
      Stack[Sp] := (Right[U] = A);
      Inc(Sp);
      A := U;
    until A = Root;

    {Unstack to transmit bits in correct order}
    repeat
      Dec(Sp);
      if Stack[Sp] then
        OutByte := OutByte or BitMask[BitPos];
      if BitPos = 7 then begin
        {Byte filled with bits, write it out}
        WriteByte;
        BitPos := 0;
        OutByte := 0;
      end else
        Inc(BitPos);
    until Sp = 0;

    {Update the tree}
    Splay(Plain);
  end;

  procedure CompressFile;
    {-Compress Inf, writing to OutF}
  var
    Header : FileHeader;
  begin
    {Write header to output}
    Header.Signature := Sig;
    BlockWrite(OutF, Header, SizeOf(FileHeader));

    {Compress file}
    OutSize := 0;
    BitPos := 0;
    OutByte := 0;
    repeat
      BlockRead(InF, InBuffer, BufSize, InSize);
      for Index := 1 to InSize do
        Compress(InBuffer[Index]);
    until InSize < BufSize;

    {Mark end of file}
    Compress(EofChar);

    {Flush buffers}
    if BitPos <> 0 then
      WriteByte;
    FlushOutBuffer;
  end;

  procedure ReadHeader;
    {-Read a compressed file header}
  var
    Header : FileHeader;
  begin
    BlockRead(InF, Header, SizeOf(FileHeader));
    if Header.Signature <> Sig then begin
      WriteLn('Unrecognized file format');
      Halt(1);
    end;
  end;

  function GetByte : Byte;
    {-Return next byte from compressed input}
  begin
    Inc(Index);
    if Index > InSize then begin
      {Reload file buffer}
      BlockRead(InF, InBuffer, BufSize, InSize);
      Index := 1;
      {End of file handled by special marker in compressed file}
    end;
    {Get next byte from buffer}
    GetByte := InBuffer[Index];
  end;

  function Expand : CodeType;
    {-Return next character from compressed input}
  var
    A : DownIndex;
  begin
    {Scan the tree to a leaf, which determines the character}
    A := Root;
    repeat
      if BitPos = 7 then begin
        {Used up the bits in current byte, get another}
        InByte := GetByte;
        BitPos := 0;
      end else
        Inc(BitPos);
      if InByte and BitMask[BitPos] = 0 then
        A := Left[A]
      else
        A := Right[A];
    until A > PredMax;

    {Update the code tree}
    Dec(A, MaxChar);
    Splay(A);

    {Return the character}
    Expand := A;
  end;

  procedure ExpandFile;
    {-Uncompress the input file and write output}
  begin
    {Force buffer load first time}
    Index := 0;
    InSize := 0;
    {Nothing in output buffer}
    OutSize := 0;
    {Force bit buffer load first time}
    BitPos := 7;

    {Read and expand the compressed input}
    OutByte := Expand;
    while OutByte <> EofChar do begin
      WriteByte;
      OutByte := Expand;
    end;

    {Flush the output}
    FlushOutBuffer;
  end;

  procedure GetParameters;
    {-Interpret command line parameters}
  var
    Arg : string[127];
  begin
    InName := '';
    OutName := '';
    CompressFlag := True;

    if ParamCount < 2 then begin
      Write('Input file : ');
      ReadLn(InName);
      Write('Output file: ');
      ReadLn(OutName);
      Write('Expand? (Y/N) ');
      ReadLn(CompStr);
      if (Length(CompStr) = 1) and (Upcase(CompStr[1]) = 'Y') then
        CompressFlag := False;
    end else
      for Index := 1 to ParamCount do begin
        Arg := ParamStr(Index);
        if (Arg[1] = '/') and (Length(Arg) = 2) then
          case Upcase(Arg[2]) of
            'X' : CompressFlag := False;
          else
            WriteLn('Unknown option: ', Arg);
            Halt(1);
          end
        else if Length(InName) = 0 then
          InName := Arg
        else if Length(OutName) = 0 then
          OutName := Arg
        else begin
          WriteLn('Too many filenames');
          Halt(1);
        end;
      end;

    if Length(InName) = 0 then
      Halt;
    if Length(OutName) = 0 then
      Halt;
  end;

begin
  GetParameters;
  InitializeSplay;

  Assign(InF, InName);
  Reset(InF, 1);
  Assign(OutF, OutName);
  Rewrite(OutF, 1);

  if CompressFlag then
    CompressFile
  else begin
    ReadHeader;
    ExpandFile;
  end;

  Close(InF);
  Close(OutF);
end.
