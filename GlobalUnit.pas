unit GlobalUnit;

interface

uses
  Classes, ToolsAPI;

type
  TLineType = (ltNon, ltClassDef, ltMethod, ltIniOrFinal, ltLogicStart, ltEnd);

  TLineInfo = record
    LineType:   TLineType;
    StartRow:   Integer;
    LogicStart: Integer;
    EndRow:     Integer;
  end;

  TLineInfoArray = array of TLineInfo;

  TMethodLineInfo = record
    CommentStart: Integer;
    StartRow:     Integer;
    LogicStart:   Integer;
    EndRow:       Integer;
  end;

function WideStringToUtf8(const S: WideString): UTF8String;
function Utf8ToWideString(const S: UTF8String): WideString;
function CheckEditView(): Boolean;
function GetAllText(): WideString; overLoad;
function GetAllText(EditBuffer: IOTAEditBuffer): WideString; overLoad;
function GetLineText(LineNo: Integer): WideString;
function GetCursorWord(): string;
function GetEOLColumn(): Integer;
function GetActLineBookMarkID(): Integer;
function GetUsedBookMarkCount(): Integer;
function GetUnusedBookMarkID(): Integer;
function GetNextBookmarkID(PrevOrNext: Integer): Integer;
procedure SelectRows(StarRow, EndRow: Integer);
function GetCommentStart(LineNo: Integer; AllText: string): Integer;
function GetMethodLineInfo(LineNo: Integer; AllText: string): TMethodLineInfo;
function GetStartLineArray(const AllText: string; out StartLineArray: TLineInfoArray): Integer;
procedure BackupSource(BackupFolder: string; EditBuffer: IOTAEditBuffer);
function GetCurrentProject(): IOTAProject;
function GetModuleFromFileName(const FileName: string): IOTAModule;

var
  EditView: IOTAEditView;
  RegExp, Match: OleVariant;

const
  C_CLASS  = '^\s+(\w+)\s*=\s*(class(|\s*\(.+))\s*$';
  C_METHOD = '^(procedure|function|constructor|destructor)\s+(((\w+)\.(\w+)|(\w+)).*)$';

implementation

uses
  Windows, SysUtils, ComObj, Dialogs, StrUtils, Math;

//*****************************************************************************
//[ �T  �v ]�@WideString��UTF8�̕�����ɕϊ�����
//[ ��  �� ]�@WideString
//[ �߂�l ]�@UTF8�̕�����
//*****************************************************************************
function WideStringToUtf8(const S: WideString): UTF8String;
var
  Len: Integer;
  Buf: array of AnsiChar;
begin
  Result := '';
  Len := WideCharToMultiByte(CP_UTF8, 0, PWideChar(S), -1, nil, 0, nil, nil);
  if Len = 0 then Exit;

  SetLength(Buf,Len);
  WideCharToMultiByte(CP_UTF8, 0, PWideChar(S), -1, PAnsiChar(Buf), Len, nil, nil);
  SetString(Result, PChar(Buf), Len);
end;

//*****************************************************************************
//[ �T  �v ]�@UTF8�̕������WideString�ɕϊ�����
//[ ��  �� ]�@UTF8�̕�����
//[ �߂�l ]�@WideString
//*****************************************************************************
function Utf8ToWideString(const S: UTF8String): WideString;
var
  Len: Integer;
  Buf: array of WideChar;
begin
  Result := '';
  Len := MultiByteToWideChar(CP_UTF8, 0, PAnsiChar(S), Length(S), nil, 0);
  if Len = 0 then Exit;

  SetLength(Buf,Len + 1); //NULL�I�[���P�����]�v�Ɋm�ۂ���
  MultiByteToWideChar(CP_UTF8, 0, PAnsiChar(S), Length(S), PWideChar(Buf), Len);
  Result := PWideChar(Buf);
end;

//*****************************************************************************
//[ �T  �v ]�@�ҏW����EditView�I�u�W�F�N�g��Private�ϐ��ɐݒ�
//[ �߂�l ]�@True:����
//*****************************************************************************
function CheckEditView(): Boolean;
begin
try
  Result := False;
  if (BorlandIDEServices as IOTAModuleServices).ModuleCount = 0 then Exit;
  EditView := (BorlandIDEServices as IOTAEditorServices).TopView;
  Result := Assigned(EditView);
except
  Result := False;
end;
end;

//*****************************************************************************
//[ �T  �v ]�@�ҏW���̃G�f�B�^�̑S��������擾
//[ �߂�l ]�@�ҏW���̃G�f�B�^�̑S������
//*****************************************************************************
function GetAllText(): WideString;
begin
  Result := GetAllText(EditView.Buffer);
end;

//*****************************************************************************
//[ �T  �v ]�@EditBuffer�̑S��������擾
//[ ��  �� ]�@EditBuffer
//[ �߂�l ]�@EditBuffer�̑S������
//*****************************************************************************
function GetAllText(EditBuffer: IOTAEditBuffer): WideString;
var
  CharPos: TOTACharPos;
  EOFPos: Integer;
  str: AnsiString;
  Reader: IOTAEditReader;
begin
  Result := '';

  //�ŏI�s�̖�����Pos���擾
  CharPos.Line := EditBuffer.EditPosition.LastRow;
  CharPos.CharIndex := 9999;
  EOFPos := EditBuffer.TopView.CharPosToPos(CharPos);

  SetLength(str, EOFPos);
  Reader := EditBuffer.CreateReader;
  Reader.GetText(0, PAnsiChar(str), EOFPos);
  Result := Utf8ToWideString(str);
end;

//*****************************************************************************
//[ �T  �v ]�@LineNo�s�̕�������擾
//[ �߂�l ]�@������
//*****************************************************************************
function GetLineText(LineNo: Integer): WideString;
var
  CharPos: TOTACharPos;
  BOLPos, EOLPos: Integer;
  str: AnsiString;
  Reader: IOTAEditReader;
begin
  //�����ƕ�����Pos���擾
  CharPos.Line := LineNo;
  CharPos.CharIndex := 0;
  BOLPos := EditView.CharPosToPos(CharPos);
  CharPos.CharIndex := 9999;
  EOLPos := EditView.CharPosToPos(CharPos);

  SetLength(str, EOLPos - BOLPos);
  Reader := EditView.Buffer.CreateReader;
  Reader.GetText(BOLPos, PAnsiChar(str), EOLPos - BOLPos);
  Result := Utf8ToWideString(str);
  //�����̉��s���폜
  Result := TrimRight(Result);
end;

//*****************************************************************************
//[ �T  �v ]�@�J�[�\���ʒu�̒P����擾
//[ �߂�l ]�@Word
//*****************************************************************************
function GetCursorWord(): string;
begin
  Result := '';

  //��`�I�����ƑΏۊO
  if EditView.Block.Style = btColumn then Exit;

  //�I���������s���ƑΏۊO
  if EditView.Block.StartingRow <> EditView.Block.EndingRow then Exit;
  //�I��͈͂̕�������擾
  if EditView.Block.Size <> 0 then
  begin
    Result := Utf8ToWideString(EditView.Block.Text);
    Exit;
  end;

  //�J�[�\�����̒P����擾
  if EditView.Position.IsWordCharacter then
  begin
    EditView.Position.Save;
    EditView.Position.MoveCursor(mmSkipLeft or mmSkipWord);
    EditView.Block.Reset;
    EditView.Block.Style := btNonInclusive;
    EditView.Block.BeginBlock;
    EditView.Position.MoveCursor(mmSkipRight or mmSkipWord);
    EditView.Block.EndBlock;
    Result := Utf8ToWideString(EditView.Block.Text);
    EditView.Position.Restore;
  end;
end;

//*****************************************************************************
//[ �T  �v ]�@�J�[�\���s�̍s���̈ʒu���擾
//*****************************************************************************
function GetEOLColumn(): Integer;
var
  EOLPos: Integer;
  CharPos: TOTACharPos;
  EditPos: TOTAEditPos;
begin
  //������Pos���擾
  CharPos.Line := EditView.Position.Row;
  CharPos.CharIndex := 9999;
  EOLPos := EditView.CharPosToPos(CharPos);
  CharPos := EditView.PosToCharPos(EOLPos - 1); //���s�͑ΏۊO

  //CharPos �� EditPos
  EditView.ConvertPos(False, EditPos, CharPos);
  Result := EditPos.Col;

//  //���̏�Ԃ�ۑ�
//  EditView.Position.Save;
//  LeftColumn := EditView.LeftColumn;
//
//  EditView.Position.MoveEOL;
//  Result := EditView.Position.Column;
//
//  //���̏�Ԃ𕜌�
//  EditView.SetTopLeft(EditView.TopRow, LeftColumn);
//  EditView.Position.Restore;
end;

//*****************************************************************************
//[ �T  �v ]�@�J�[�\���̂���s��BookMarkID���擾(�Ȃ��Ƃ���-1)
//[ �߂�l ]�@BookMarkID
//*****************************************************************************
function GetActLineBookMarkID(): Integer;
var
  i: Integer;
  Row: Integer;
begin
  Result := -1;
  Row := EditView.Position.Row;

  //�u�b�N�}�[�N�����[�v
  for i := 0 to 9 do
  begin
    if Row = EditView.BookmarkPos[i].Line then
    begin
      Result := i;
      Exit;
    end;
  end;
end;

//*****************************************************************************
//[ �T  �v ]�@�g�p����Ă���BookMarkID�̐����擾����
//[ �߂�l ]�@�g�p����Ă���BookMarkID�̐�
//*****************************************************************************
function GetUsedBookMarkCount(): Integer;
var
  i: Integer;
begin
  Result := 0;

  //�u�b�N�}�[�N�����[�v
  for i := 0 to 9 do
  begin
    if EditView.BookmarkPos[i].Line <> 0 then Inc(Result);
  end;
end;

//*****************************************************************************
//[ �T  �v ]�@�g�p����Ă��Ȃ��ŏ���BookMarkID���擾(���ׂĎg�p�ς݂̎���10)
//[ �߂�l ]�@BookMarkID
//*****************************************************************************
function GetUnusedBookMarkID(): Integer;
var
  i: Integer;
begin
  Result := 10;

  //�u�b�N�}�[�N�����[�v
  for i := 0 to 9 do
  begin
    if EditView.BookmarkPos[i].Line = 0 then
    begin
      Result := i;
      Exit;
    end;
  end;
end;

//*****************************************************************************
//[ �T  �v ]�@�J�[�\���s�̒��O�܂��͒����Bookmark���擾(�Ȃ��Ƃ���-1)
//[ ��  �� ]�@PrevOrNext:-1=�O�������A+1:�������
//[ �߂�l ]�@BookMarkID
//*****************************************************************************
function GetNextBookmarkID(PrevOrNext: Integer): Integer;
  function SerchNextBookmarkID(ActRow: Integer): Integer;
  var
    i: Integer;
    BookRow, Row: Integer;
  begin
    Result := -1;
    if PrevOrNext < 0 then  Row := 0
                      else  Row := EditView.Position.LastRow + 1;
    //�u�b�N�}�[�N�����[�v
    for i := 0 to 9 do
    begin
      BookRow := EditView.BookmarkPos[i].Line;
      if BookRow > 0 then
      begin
        if (PrevOrNext < 0) and (BookRow < ActRow) then
        begin
          if BookRow > Row then
          begin
            Row := BookRow;
            Result := i;
          end;
        end else
        if (PrevOrNext > 0) and (BookRow > ActRow) then
        begin
          if BookRow < Row then
          begin
            Row := BookRow;
            Result := i;
          end;
        end;
      end;
    end;
  end;
var
  ActRow: Integer;
begin
  ActRow := EditView.Position.Row;
  Result := SerchNextBookmarkID(ActRow);
  if Result >= 0 then Exit;

  //������Ȃ��������A�܂�Ԃ��čő�܂��͍ŏ���Bookmark��T��
  if PrevOrNext < 0 then  ActRow := EditView.Position.LastRow + 1
                    else  ActRow := 0;
  Result := SerchNextBookmarkID(ActRow);
end;

//*****************************************************************************
//[ �T  �v ]�@�擪�s����ŏI�s�܂ōs�I�����s��
//[ ��  �� ]�@�擪�s�A�ŏI�s
//*****************************************************************************
procedure SelectRows(StarRow, EndRow: Integer);
begin
  EditView.Position.Move(StarRow, 1);
  EditView.Block.Reset;
  EditView.Block.Style := btNonInclusive;
  EditView.Block.BeginBlock;
  if EditView.Position.LastRow > EndRow then
    EditView.Position.Move(EndRow + 1, 1)
  else
    EditView.Position.MoveEOF;
  EditView.Block.EndBlock;
end;

//*****************************************************************************
//[ �T  �v ]�@Line�s�̒��O�̃R�����g�J�n�s���擾����
//[ ��  �� ]�@�s�ԍ��A���ׂĂ̍s
//[ �߂�l ]�@TMethodLineInfo
//*****************************************************************************
function GetCommentStart(LineNo: Integer; AllText: string): Integer;
const
  MAXCOMMENT = 15;
var
  i, j: Integer;
  ST: TStringList;
begin
  Result := LineNo;

  ST := TStringList.Create;
  try
    ST.Text := AllText;
    //�s����Index�����킹�邽�ߐ擪��1�s�}��
    ST.Insert(0, '');

    //�� '//�`' or '{�`}' or '(*�`*)'
    RegExp.Pattern := '(^//.*$|^\{.*\}$|^\(\*.*\*\)$)';
    //StartRow��1�s�O����1�s���ɏ�����Ƀ��[�v���A�R�����g�̐擪�s������
    for i := LineNo - 1 downto 1 do
    begin
      if not RegExp.Test(Trim(ST[i])) then
      begin
        Result := i + 1;
        Break;
      end;
    end;

    if LineNo = Result then
    begin
      //��  '�`}' or '�`*)'
      RegExp.Pattern := '(\}|\*\))$';
      //StartRow��1�s�O���R�����g�̏I���s�̎�
      if RegExp.Test(Trim(ST[LineNo - 1])) then
      begin
        j :=0;
        //��  '{�`' or '(*�`'
        RegExp.Pattern := '^(\{|\(\*)';
        //StartRow��2�s�O����1�s���ɏ�����Ƀ��[�v���A�R�����g�̐擪�s������
        for i := LineNo - 2 downto 1 do
        begin
          Inc(j);
          //�R�����g�J�n�s�̌����ɏ����݂���
          if j > MAXCOMMENT then Break;
          if RegExp.Test(Trim(ST[i])) then
          begin
            Result := i;
            Break;
          end;
        end;
      end;
    end;
  finally
    ST.Free;
  end;

  if Result = LineNo then
    Result := LineNo - 1;
end;

//*****************************************************************************
//[ �T  �v ]�@�N���X�錾�s or ���\�b�h�J�n�setc �̈ʒu��z��Ɏ擾����
//[ ��  �� ]�@�ҏW�t�@�C���̑S�s�A
//[ ��  �� ]�@StartLineArray�F�N���X�錾�s or ���\�b�h�J�n�setc �̔z��
//[ �߂�l ]�@�N���X�錾�s or ���\�b�h�J�n�s �̐�
//*****************************************************************************
function GetStartLineArray(const AllText: string; out StartLineArray: TLineInfoArray): Integer;
var
  i, j: Integer;
  NextStartRow: Integer;
  ST: TStringList;
  TmpArray: array of TLineType;
  interfaceRow, implementationRow: Integer; //interface,implementation�̍s
begin
  StartLineArray := nil;
  Result := 0;

  ST := TStringList.Create();
  try
    ST.Text := AllText;

    //�s����Index�����킹�邽�ߐ擪��1�s�}��
    ST.Insert(0, '');

    //interface,implementation�̍s���擾
    implementationRow := 0;
    interfaceRow := 0;
    for i := 1 to ST.Count - 1 do
    begin
      if LowerCase(LeftStr(Trim(ST[i]), 9)) = 'interface' then
      begin
        interfaceRow := i;
      end;
      if LowerCase(LeftStr(Trim(ST[i]),14)) = 'implementation' then
      begin
        implementationRow := i;
        Break;
      end;
    end;

    SetLength(TmpArray, ST.Count);
    j := 0;

    //�s�������[�v���A�N���X�錾�s���擾
    RegExp.Pattern := C_CLASS;
    for i := interfaceRow + 1 to ST.Count - 1 do
    begin
      if RegExp.Test(ST[i]) then
      begin
        TmpArray[i] := ltClassDef;
        Inc(j);
      end;
    end;

    //�s�������[�v��(implementation��)�A���\�b�h�J�n�s���擾
    RegExp.Pattern := C_METHOD;
    for i := implementationRow + 1 to ST.Count - 1 do
    begin
      //forward�錾�͑ΏۊO
      if LowerCase(RightStr(Trim(ST[i]),8)) = 'forward;' then Continue;
      if RegExp.Test(ST[i]) then
      begin
        TmpArray[i] := ltMethod;
        Inc(j);
      end;
    end;

    //�ŏI�s���烋�[�v���Ainitialization��finalization���擾
    for i := ST.Count - 1 downto implementationRow + 1 do
    begin
      //���\�b�h
      if Ord(TmpArray[i]) <> 0 then Break;
      if LowerCase(LeftStr(Trim(ST[i]),14)) = 'initialization' then
      begin
        TmpArray[i] := ltIniOrFinal;
        Inc(j);
      end else
      if LowerCase(LeftStr(Trim(ST[i]),12)) = 'finalization' then
      begin
        TmpArray[i] := ltIniOrFinal;
        Inc(j);
      end;
    end;

    Result := j;
    if Result = 0 then Exit;

    SetLength(StartLineArray, j);
    j := 0;

    //�s�������[�v���AStartLineArray��StartRow��ݒ�
    for i := 1 to ST.Count - 1 do
    begin
      if Ord(TmpArray[i]) <> 0 then
      begin
        StartLineArray[j].LineType := TmpArray[i];
        StartLineArray[j].StartRow := i;
        Inc(j);
      end;
    end;

    //StartLineArray�̔z��̐��������[�v���ALogicStart,EndRow��ݒ�
    for i := Low(StartLineArray) to High(StartLineArray) do
    begin
      with StartLineArray[i] do
      begin
        //���ɐݒ�
        LogicStart := StartRow;
        EndRow     := StartRow;

        //���̃��\�b�h�̊J�n�s���擾
        if i < High(StartLineArray) then
          NextStartRow := StartLineArray[i+1].StartRow - 1
        else
          //�Ō�̃��\�b�h�̎��͍ŏI�s��ݒ�
          NextStartRow := ST.Count - 1;

        case LineType of
        ltClassDef:
        begin
          //StartRow���玟�̃��\�b�h�̊J�n�s�܂Ń��[�v
          for j := StartRow to NextStartRow do
          begin
            if LowerCase(LeftStr(Trim(ST[j]),4)) = 'end;' then
            begin
              EndRow := j;
              Break;
            end;
          end;
        end;
        ltMethod:
        begin
          //StartRow���玟�̃��\�b�h�̊J�n�s�܂Ń��[�v
          for j := StartRow to NextStartRow do
          begin
            if LowerCase(LeftStr(ST[j],5)) = 'begin' then
            begin
              LogicStart := j + 1;
              Break;
            end;
          end;
          //LogicStart���玟�̃��\�b�h�̊J�n�s�܂Ń��[�v
          for j := LogicStart to NextStartRow do
          begin
            if LowerCase(LeftStr(ST[j],4)) = 'end;' then
            begin
              EndRow := j;
              Break;
            end;
          end;
        end;
      end;
    end;
  end;
  finally
    ST.Free;
  end;
end;

//*****************************************************************************
//[ �T  �v ]�@Line���܂ރ��\�b�h�̍s�����擾
//[ ��  �� ]�@�s�ԍ��A���ׂĂ̍s
//[ �߂�l ]�@TMethodLineInfo
//*****************************************************************************
function GetMethodLineInfo(LineNo: Integer; AllText: string): TMethodLineInfo;
var
  i, j: Integer;
  ST: TStringList;
begin
  //Result�̏�����
  with Result do
  begin
    CommentStart := 0; //���\�b�h�̃R�����g�̊J�n�s
    StartRow     := 0; //�� procedure GetXXX();
    LogicStart   := 0; //�� begin �̎��̍s
    EndRow       := 0; //�� end;
  end;

  ST := TStringList.Create;
  try
    ST.Text := AllText;
    //�s����Index�����킹�邽�ߐ擪��1�s�}��
    ST.Insert(0, '');

    //implementation �s�̎擾
    j := 1;
    for i := 1 to ST.Count - 1 do
    begin
      if LowerCase(Copy(Trim(ST[i]),1,14)) = 'implementation' then
      begin
        if LineNo <= i then Exit;
        //implementation �s�ȑO�͌����̑ΏۂƂ��Ȃ�
        j := i + 1;
        Break;
      end;
    end;

    //******************************************************************
    // StartRow��ݒ�
    //******************************************************************
    RegExp.Pattern := '^(procedure|function|constructor|destructor|initialization|finalization)';
    //Line�ʒu����1�s���ɏ�����Ƀ��[�v���A���\�b�h�̐擪�s������
    for i := Min(LineNo, ST.Count - 1) downto j do
    begin
      if RegExp.Test(ST[i]) then
      begin
        Result.StartRow := i;
        Break;
      end;
      //�O���̃��\�b�h�̏I���������Ă��܂���
      if LowerCase(Copy(ST[i - 1],1,4)) = 'end;' then Exit;
    end;
    //������Ȃ�������
    if Result.StartRow = 0 then Exit;

    //******************************************************************
    // CommentStart��ݒ�
    //******************************************************************
    Result.CommentStart := GetCommentStart(Result.StartRow, AllText);

//    //�� '//�`' or '{�`}' or '(*�`*)'
//    RegExp.Pattern := '(^//.*$|^\{.*\}$|^\(\*.*\*\)$)';
//    //StartRow��1�s�O����1�s���ɏ�����Ƀ��[�v���A�R�����g�̐擪�s������
//    for i := Result.StartRow - 1 downto j do
//    begin
//      if not RegExp.Test(Trim(ST[i])) then
//      begin
//        Result.CommentStart := i + 1;
//        Break;
//      end;
//    end;
//
//    if Result.CommentStart = Result.StartRow then
//    begin
//      //��  '�`}' or '�`*)'
//      RegExp.Pattern := '(\}|\*\))$';
//      //StartRow��1�s�O���R�����g�̏I���s�̎�
//      if RegExp.Test(Trim(ST[Result.StartRow - 1])) then
//      begin
//        k :=0;
//        //��  '{�`' or '(*�`'
//        RegExp.Pattern := '^(\{|\(\*)';
//        //StartRow��2�s�O����1�s���ɏ�����Ƀ��[�v���A�R�����g�̐擪�s������
//        for i := Result.StartRow - 2 downto j do
//        begin
//          Inc(k);
//          //�R�����g�J�n�s�̌����ɏ����݂���
//          if k > MAXCOMMENT then Break;
//          if RegExp.Test(Trim(ST[i])) then
//          begin
//            Result.CommentStart := i;
//            Break;
//          end;
//        end;
//      end;
//    end;

    //******************************************************************
    // LogicStart��ݒ�
    //******************************************************************
    RegExp.Pattern := '^(procedure|function|constructor|destructor|initialization|finalization)';
    //StartRow��1�s�ォ��1�s���ɉ������Ƀ��[�v���A���W�b�N�̐擪�s������
    for i := Result.StartRow + 1 to ST.Count - 2 do  //��O�ƂȂ�Ȃ��悤��2������
    begin
      if LowerCase(Copy(ST[i],1,5)) = 'begin' then
      begin
        Result.LogicStart := i + 1;
        Break;
      end;

      //���̃��\�b�h�ɂȂ��Ă��܂����B
      if RegExp.Test(ST[i]) then
      begin
        Result.LogicStart := Result.StartRow;
        Result.EndRow := Result.StartRow;
        Exit;
      end;
    end;
    //������Ȃ�������
    if Result.LogicStart = 0 then
    begin
      Result.LogicStart := Result.StartRow;
      Result.EndRow := Result.StartRow;
      Exit;
    end;

    //******************************************************************
    // EndRow��ݒ�
    //******************************************************************
    RegExp.Pattern := '^(procedure|function|constructor|destructor|initialization|finalization)';
    //LogicStart��1�s�ォ��1�s���ɉ������Ƀ��[�v���A���\�b�h�̏I���s������
    for i := Result.LogicStart + 1 to ST.Count - 1 do
    begin
      if LowerCase(Copy(ST[i],1,4)) = 'end;' then
      begin
        Result.EndRow := i;
        Break;
      end;

      //���̃��\�b�h�ɂȂ��Ă��܂����B
      if RegExp.Test(ST[i]) then
      begin
        Result.LogicStart := Result.StartRow;
        Result.EndRow := Result.StartRow;
        Exit;
      end;
    end;
    //������Ȃ�������
    if Result.EndRow = 0 then
    begin
      Result.EndRow := Result.LogicStart;
      Exit;
    end;

  finally
    ST.Free;
  end;
end;

//*****************************************************************************
//[ �T  �v ]�@BackupFolder��EditBuffer�̃\�[�X��ۑ�����
//[ ��  �� ]�@BackupFolder����EditBuffer
//*****************************************************************************
procedure BackupSource(BackupFolder: string; EditBuffer: IOTAEditBuffer);
var
  FileName: string;
  ST: TStringList;
begin
  FileName := BackupFolder + '\' + ExtractFileName(EditBuffer.FileName);

  //�o�b�N�A�b�v�t�H���_�����݂��Ȃ���΍쐬
  if not DirectoryExists(BackupFolder) then
  begin
    if not CreateDir(BackupFolder) then Exit;
  end;

  //�\�[�X��ۑ�
  ST := TStringList.Create;
  try try
    ST.Text := GetAllText(EditBuffer);
    ST.SaveToFile(FileName);
  finally
    ST.Free;
  end;
  except end;
end;
//*****************************************************************************
//[ �T  �v ]�@�J�����g��IOTAProject�C���^�[�t�F�[�X�I�u�W�F�N�g���擾
//[ �߂�l ]�@IOTAProject
//*****************************************************************************
function GetCurrentProject(): IOTAProject;
var
  i: Integer;
  ProjectGroup: IOTAProjectGroup;
  ModuleServices: IOTAModuleServices;
  Module: IOTAModule;
begin
  Result := nil;

  ModuleServices := (BorlandIDEServices as IOTAModuleServices);
  if ModuleServices.ModuleCount = 0 then Exit;

  //Module�̐��������[�v
  for i := 0 to ModuleServices.ModuleCount - 1 do
  begin
    Module := ModuleServices.Modules[i];
    if Supports(Module, IOTAProjectGroup, ProjectGroup) then
    begin
      Result := ProjectGroup.ActiveProject;
      Exit;
    end;
  end;

  //Module�̐��������[�v
  for i := 0 to ModuleServices.ModuleCount - 1 do
  begin
    Module := ModuleServices.Modules[i];
    if Supports(Module, IOTAProject) then
    begin
      Supports(Module, IOTAProject, Result);
      Exit;
    end;
  end;
end;

//*****************************************************************************
//[ �T  �v ]�@�t�@�C��������IOTAModule�C���^�[�t�F�[�X�I�u�W�F�N�g���擾
//[ ��  �� ]�@�t�@�C����
//[ �߂�l ]�@IOTAModule
//*****************************************************************************
function GetModuleFromFileName(const FileName: string): IOTAModule;
var
  ModuleServices: IOTAModuleServices;
begin
  Result := nil;

  ModuleServices := (BorlandIDEServices as IOTAModuleServices);
  if ModuleServices.ModuleCount = 0 then Exit;
  Result := ModuleServices.FindModule(FileName);
end;

//*****************************************************************************
//initialization
//*****************************************************************************
initialization
  //�������̂��߂����Ő��K�\���̏���
  RegExp := CreateOleObject('VBScript.RegExp');
  RegExp.IgnoreCase := True;  //�啶���Ə���������ʂ��Ȃ�

//*****************************************************************************
//finalization
//*****************************************************************************
finalization
end.
