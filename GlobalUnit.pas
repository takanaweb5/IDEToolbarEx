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
//[ 概  要 ]　WideStringをUTF8の文字列に変換する
//[ 引  数 ]　WideString
//[ 戻り値 ]　UTF8の文字列
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
//[ 概  要 ]　UTF8の文字列をWideStringに変換する
//[ 引  数 ]　UTF8の文字列
//[ 戻り値 ]　WideString
//*****************************************************************************
function Utf8ToWideString(const S: UTF8String): WideString;
var
  Len: Integer;
  Buf: array of WideChar;
begin
  Result := '';
  Len := MultiByteToWideChar(CP_UTF8, 0, PAnsiChar(S), Length(S), nil, 0);
  if Len = 0 then Exit;

  SetLength(Buf,Len + 1); //NULL終端分１文字余計に確保する
  MultiByteToWideChar(CP_UTF8, 0, PAnsiChar(S), Length(S), PWideChar(Buf), Len);
  Result := PWideChar(Buf);
end;

//*****************************************************************************
//[ 概  要 ]　編集中のEditViewオブジェクトをPrivate変数に設定
//[ 戻り値 ]　True:成功
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
//[ 概  要 ]　編集中のエディタの全文字列を取得
//[ 戻り値 ]　編集中のエディタの全文字列
//*****************************************************************************
function GetAllText(): WideString;
begin
  Result := GetAllText(EditView.Buffer);
end;

//*****************************************************************************
//[ 概  要 ]　EditBufferの全文字列を取得
//[ 引  数 ]　EditBuffer
//[ 戻り値 ]　EditBufferの全文字列
//*****************************************************************************
function GetAllText(EditBuffer: IOTAEditBuffer): WideString;
var
  CharPos: TOTACharPos;
  EOFPos: Integer;
  str: AnsiString;
  Reader: IOTAEditReader;
begin
  Result := '';

  //最終行の末尾のPosを取得
  CharPos.Line := EditBuffer.EditPosition.LastRow;
  CharPos.CharIndex := 9999;
  EOFPos := EditBuffer.TopView.CharPosToPos(CharPos);

  SetLength(str, EOFPos);
  Reader := EditBuffer.CreateReader;
  Reader.GetText(0, PAnsiChar(str), EOFPos);
  Result := Utf8ToWideString(str);
end;

//*****************************************************************************
//[ 概  要 ]　LineNo行の文字列を取得
//[ 戻り値 ]　文字列
//*****************************************************************************
function GetLineText(LineNo: Integer): WideString;
var
  CharPos: TOTACharPos;
  BOLPos, EOLPos: Integer;
  str: AnsiString;
  Reader: IOTAEditReader;
begin
  //文頭と文末のPosを取得
  CharPos.Line := LineNo;
  CharPos.CharIndex := 0;
  BOLPos := EditView.CharPosToPos(CharPos);
  CharPos.CharIndex := 9999;
  EOLPos := EditView.CharPosToPos(CharPos);

  SetLength(str, EOLPos - BOLPos);
  Reader := EditView.Buffer.CreateReader;
  Reader.GetText(BOLPos, PAnsiChar(str), EOLPos - BOLPos);
  Result := Utf8ToWideString(str);
  //末尾の改行を削除
  Result := TrimRight(Result);
end;

//*****************************************************************************
//[ 概  要 ]　カーソル位置の単語を取得
//[ 戻り値 ]　Word
//*****************************************************************************
function GetCursorWord(): string;
begin
  Result := '';

  //矩形選択だと対象外
  if EditView.Block.Style = btColumn then Exit;

  //選択が複数行だと対象外
  if EditView.Block.StartingRow <> EditView.Block.EndingRow then Exit;
  //選択範囲の文字列を取得
  if EditView.Block.Size <> 0 then
  begin
    Result := Utf8ToWideString(EditView.Block.Text);
    Exit;
  end;

  //カーソル下の単語を取得
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
//[ 概  要 ]　カーソル行の行末の位置を取得
//*****************************************************************************
function GetEOLColumn(): Integer;
var
  EOLPos: Integer;
  CharPos: TOTACharPos;
  EditPos: TOTAEditPos;
begin
  //文末のPosを取得
  CharPos.Line := EditView.Position.Row;
  CharPos.CharIndex := 9999;
  EOLPos := EditView.CharPosToPos(CharPos);
  CharPos := EditView.PosToCharPos(EOLPos - 1); //改行は対象外

  //CharPos → EditPos
  EditView.ConvertPos(False, EditPos, CharPos);
  Result := EditPos.Col;

//  //元の状態を保存
//  EditView.Position.Save;
//  LeftColumn := EditView.LeftColumn;
//
//  EditView.Position.MoveEOL;
//  Result := EditView.Position.Column;
//
//  //元の状態を復元
//  EditView.SetTopLeft(EditView.TopRow, LeftColumn);
//  EditView.Position.Restore;
end;

//*****************************************************************************
//[ 概  要 ]　カーソルのある行のBookMarkIDを取得(ないときは-1)
//[ 戻り値 ]　BookMarkID
//*****************************************************************************
function GetActLineBookMarkID(): Integer;
var
  i: Integer;
  Row: Integer;
begin
  Result := -1;
  Row := EditView.Position.Row;

  //ブックマーク数ループ
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
//[ 概  要 ]　使用されているBookMarkIDの数を取得する
//[ 戻り値 ]　使用されているBookMarkIDの数
//*****************************************************************************
function GetUsedBookMarkCount(): Integer;
var
  i: Integer;
begin
  Result := 0;

  //ブックマーク数ループ
  for i := 0 to 9 do
  begin
    if EditView.BookmarkPos[i].Line <> 0 then Inc(Result);
  end;
end;

//*****************************************************************************
//[ 概  要 ]　使用されていない最小のBookMarkIDを取得(すべて使用済みの時は10)
//[ 戻り値 ]　BookMarkID
//*****************************************************************************
function GetUnusedBookMarkID(): Integer;
var
  i: Integer;
begin
  Result := 10;

  //ブックマーク数ループ
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
//[ 概  要 ]　カーソル行の直前または直後のBookmarkを取得(ないときは-1)
//[ 引  数 ]　PrevOrNext:-1=前方検索、+1:後方検索
//[ 戻り値 ]　BookMarkID
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
    //ブックマーク数ループ
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

  //見つからなかった時、折り返して最大または最小のBookmarkを探す
  if PrevOrNext < 0 then  ActRow := EditView.Position.LastRow + 1
                    else  ActRow := 0;
  Result := SerchNextBookmarkID(ActRow);
end;

//*****************************************************************************
//[ 概  要 ]　先頭行から最終行まで行選択を行う
//[ 引  数 ]　先頭行、最終行
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
//[ 概  要 ]　Line行の直前のコメント開始行を取得する
//[ 引  数 ]　行番号、すべての行
//[ 戻り値 ]　TMethodLineInfo
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
    //行数とIndexをあわせるため先頭に1行挿入
    ST.Insert(0, '');

    //例 '//～' or '{～}' or '(*～*)'
    RegExp.Pattern := '(^//.*$|^\{.*\}$|^\(\*.*\*\)$)';
    //StartRowの1行前から1行毎に上方向にループし、コメントの先頭行を検索
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
      //例  '～}' or '～*)'
      RegExp.Pattern := '(\}|\*\))$';
      //StartRowの1行前がコメントの終了行の時
      if RegExp.Test(Trim(ST[LineNo - 1])) then
      begin
        j :=0;
        //例  '{～' or '(*～'
        RegExp.Pattern := '^(\{|\(\*)';
        //StartRowの2行前から1行毎に上方向にループし、コメントの先頭行を検索
        for i := LineNo - 2 downto 1 do
        begin
          Inc(j);
          //コメント開始行の検索に上限を設ける
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
//[ 概  要 ]　クラス宣言行 or メソッド開始行etc の位置を配列に取得する
//[ 引  数 ]　編集ファイルの全行、
//[ 結  果 ]　StartLineArray：クラス宣言行 or メソッド開始行etc の配列
//[ 戻り値 ]　クラス宣言行 or メソッド開始行 の数
//*****************************************************************************
function GetStartLineArray(const AllText: string; out StartLineArray: TLineInfoArray): Integer;
var
  i, j: Integer;
  NextStartRow: Integer;
  ST: TStringList;
  TmpArray: array of TLineType;
  interfaceRow, implementationRow: Integer; //interface,implementationの行
begin
  StartLineArray := nil;
  Result := 0;

  ST := TStringList.Create();
  try
    ST.Text := AllText;

    //行数とIndexをあわせるため先頭に1行挿入
    ST.Insert(0, '');

    //interface,implementationの行を取得
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

    //行数分ループし、クラス宣言行を取得
    RegExp.Pattern := C_CLASS;
    for i := interfaceRow + 1 to ST.Count - 1 do
    begin
      if RegExp.Test(ST[i]) then
      begin
        TmpArray[i] := ltClassDef;
        Inc(j);
      end;
    end;

    //行数分ループし(implementation節)、メソッド開始行を取得
    RegExp.Pattern := C_METHOD;
    for i := implementationRow + 1 to ST.Count - 1 do
    begin
      //forward宣言は対象外
      if LowerCase(RightStr(Trim(ST[i]),8)) = 'forward;' then Continue;
      if RegExp.Test(ST[i]) then
      begin
        TmpArray[i] := ltMethod;
        Inc(j);
      end;
    end;

    //最終行からループし、initializationとfinalizationを取得
    for i := ST.Count - 1 downto implementationRow + 1 do
    begin
      //メソッド
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

    //行数分ループし、StartLineArrayのStartRowを設定
    for i := 1 to ST.Count - 1 do
    begin
      if Ord(TmpArray[i]) <> 0 then
      begin
        StartLineArray[j].LineType := TmpArray[i];
        StartLineArray[j].StartRow := i;
        Inc(j);
      end;
    end;

    //StartLineArrayの配列の数だけループし、LogicStart,EndRowを設定
    for i := Low(StartLineArray) to High(StartLineArray) do
    begin
      with StartLineArray[i] do
      begin
        //仮に設定
        LogicStart := StartRow;
        EndRow     := StartRow;

        //次のメソッドの開始行を取得
        if i < High(StartLineArray) then
          NextStartRow := StartLineArray[i+1].StartRow - 1
        else
          //最後のメソッドの時は最終行を設定
          NextStartRow := ST.Count - 1;

        case LineType of
        ltClassDef:
        begin
          //StartRowから次のメソッドの開始行までループ
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
          //StartRowから次のメソッドの開始行までループ
          for j := StartRow to NextStartRow do
          begin
            if LowerCase(LeftStr(ST[j],5)) = 'begin' then
            begin
              LogicStart := j + 1;
              Break;
            end;
          end;
          //LogicStartから次のメソッドの開始行までループ
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
//[ 概  要 ]　Lineを含むメソッドの行情報を取得
//[ 引  数 ]　行番号、すべての行
//[ 戻り値 ]　TMethodLineInfo
//*****************************************************************************
function GetMethodLineInfo(LineNo: Integer; AllText: string): TMethodLineInfo;
var
  i, j: Integer;
  ST: TStringList;
begin
  //Resultの初期化
  with Result do
  begin
    CommentStart := 0; //メソッドのコメントの開始行
    StartRow     := 0; //例 procedure GetXXX();
    LogicStart   := 0; //例 begin の次の行
    EndRow       := 0; //例 end;
  end;

  ST := TStringList.Create;
  try
    ST.Text := AllText;
    //行数とIndexをあわせるため先頭に1行挿入
    ST.Insert(0, '');

    //implementation 行の取得
    j := 1;
    for i := 1 to ST.Count - 1 do
    begin
      if LowerCase(Copy(Trim(ST[i]),1,14)) = 'implementation' then
      begin
        if LineNo <= i then Exit;
        //implementation 行以前は検索の対象としない
        j := i + 1;
        Break;
      end;
    end;

    //******************************************************************
    // StartRowを設定
    //******************************************************************
    RegExp.Pattern := '^(procedure|function|constructor|destructor|initialization|finalization)';
    //Line位置から1行毎に上方向にループし、メソッドの先頭行を検索
    for i := Min(LineNo, ST.Count - 1) downto j do
    begin
      if RegExp.Test(ST[i]) then
      begin
        Result.StartRow := i;
        Break;
      end;
      //前方のメソッドの終了を見つけてしまった
      if LowerCase(Copy(ST[i - 1],1,4)) = 'end;' then Exit;
    end;
    //見つからなかった時
    if Result.StartRow = 0 then Exit;

    //******************************************************************
    // CommentStartを設定
    //******************************************************************
    Result.CommentStart := GetCommentStart(Result.StartRow, AllText);

//    //例 '//～' or '{～}' or '(*～*)'
//    RegExp.Pattern := '(^//.*$|^\{.*\}$|^\(\*.*\*\)$)';
//    //StartRowの1行前から1行毎に上方向にループし、コメントの先頭行を検索
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
//      //例  '～}' or '～*)'
//      RegExp.Pattern := '(\}|\*\))$';
//      //StartRowの1行前がコメントの終了行の時
//      if RegExp.Test(Trim(ST[Result.StartRow - 1])) then
//      begin
//        k :=0;
//        //例  '{～' or '(*～'
//        RegExp.Pattern := '^(\{|\(\*)';
//        //StartRowの2行前から1行毎に上方向にループし、コメントの先頭行を検索
//        for i := Result.StartRow - 2 downto j do
//        begin
//          Inc(k);
//          //コメント開始行の検索に上限を設ける
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
    // LogicStartを設定
    //******************************************************************
    RegExp.Pattern := '^(procedure|function|constructor|destructor|initialization|finalization)';
    //StartRowの1行後から1行毎に下方向にループし、ロジックの先頭行を検索
    for i := Result.StartRow + 1 to ST.Count - 2 do  //例外とならないように2を引く
    begin
      if LowerCase(Copy(ST[i],1,5)) = 'begin' then
      begin
        Result.LogicStart := i + 1;
        Break;
      end;

      //次のメソッドになってしまった。
      if RegExp.Test(ST[i]) then
      begin
        Result.LogicStart := Result.StartRow;
        Result.EndRow := Result.StartRow;
        Exit;
      end;
    end;
    //見つからなかった時
    if Result.LogicStart = 0 then
    begin
      Result.LogicStart := Result.StartRow;
      Result.EndRow := Result.StartRow;
      Exit;
    end;

    //******************************************************************
    // EndRowを設定
    //******************************************************************
    RegExp.Pattern := '^(procedure|function|constructor|destructor|initialization|finalization)';
    //LogicStartの1行後から1行毎に下方向にループし、メソッドの終了行を検索
    for i := Result.LogicStart + 1 to ST.Count - 1 do
    begin
      if LowerCase(Copy(ST[i],1,4)) = 'end;' then
      begin
        Result.EndRow := i;
        Break;
      end;

      //次のメソッドになってしまった。
      if RegExp.Test(ST[i]) then
      begin
        Result.LogicStart := Result.StartRow;
        Result.EndRow := Result.StartRow;
        Exit;
      end;
    end;
    //見つからなかった時
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
//[ 概  要 ]　BackupFolderにEditBufferのソースを保存する
//[ 引  数 ]　BackupFolder名とEditBuffer
//*****************************************************************************
procedure BackupSource(BackupFolder: string; EditBuffer: IOTAEditBuffer);
var
  FileName: string;
  ST: TStringList;
begin
  FileName := BackupFolder + '\' + ExtractFileName(EditBuffer.FileName);

  //バックアップフォルダが存在しなければ作成
  if not DirectoryExists(BackupFolder) then
  begin
    if not CreateDir(BackupFolder) then Exit;
  end;

  //ソースを保存
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
//[ 概  要 ]　カレントのIOTAProjectインターフェースオブジェクトを取得
//[ 戻り値 ]　IOTAProject
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

  //Moduleの数だけループ
  for i := 0 to ModuleServices.ModuleCount - 1 do
  begin
    Module := ModuleServices.Modules[i];
    if Supports(Module, IOTAProjectGroup, ProjectGroup) then
    begin
      Result := ProjectGroup.ActiveProject;
      Exit;
    end;
  end;

  //Moduleの数だけループ
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
//[ 概  要 ]　ファイル名からIOTAModuleインターフェースオブジェクトを取得
//[ 引  数 ]　ファイル名
//[ 戻り値 ]　IOTAModule
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
  //高速化のためここで正規表現の準備
  RegExp := CreateOleObject('VBScript.RegExp');
  RegExp.IgnoreCase := True;  //大文字と小文字を区別しない

//*****************************************************************************
//finalization
//*****************************************************************************
finalization
end.
