unit KeyBinding;

interface

procedure Register;

implementation

uses Windows, Classes, SysUtils, ToolsAPI, StrUtils, Menus, Module, GlobalUnit,
     MethodForm, Dialogs, Messages;

type
	TKeyBinding = class(TNotifierObject, IOTAKeyboardBinding)
  private
    FAllText: string;
    FStartLineArray: TLineInfoArray;
		procedure GoNextMethod (const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
		procedure SelectMethod (const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
    procedure LeftKeyPress (const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
    procedure RightKeyPress(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
    function WaitForKeyMsgOver(): Boolean;
    procedure JumpNextMethod(PrevOrNext: Integer);
	public
		{お約束}
		procedure BindKeyboard(const BindingServices: IOTAKeyBindingServices);
		function GetBindingType: TBindingType;
		function GetDisplayName: string;
		function GetName: string;
	end;

var
	BindNo: Integer;

{ TKeyBinding }
//*****************************************************************************
//[ 概  要 ]　キーの登録
//*****************************************************************************
procedure TKeyBinding.BindKeyboard(const BindingServices: IOTAKeyBindingServices);
begin
	//Ctrl + ↓
	BindingServices.AddKeyBinding([ShortCut(VK_DOWN,[ssCtrl])], GoNextMethod, Pointer(+1));
	//Ctrl + ↑
	BindingServices.AddKeyBinding([ShortCut(VK_UP,  [ssCtrl])], GoNextMethod, Pointer(-1));

	//←
	BindingServices.AddKeyBinding([ShortCut(VK_LEFT, [])], LeftKeyPress, nil);
	//→
	BindingServices.AddKeyBinding([ShortCut(VK_RIGHT,[])], RightKeyPress, nil);

	//Ctrl + Shitf + A
	BindingServices.AddKeyBinding([ShortCut(Ord('A'), [ssCtrl, ssShift])], SelectMethod, nil);

//	//Tab
//	BindingServices.AddKeyBinding([ShortCut(VK_TAB, [])],       TabKeyPress, Pointer(+1));
//	//Shift + Tab
//	BindingServices.AddKeyBinding([ShortCut(VK_TAB, [ssShift])],TabKeyPress, Pointer(-1));
//
//	//F10
//	BindingServices.AddKeyBinding([ShortCut(VK_F10, [])], ShowForm, nil);
end;

//*****************************************************************************
//[ 概  要 ]　次のメソッドへジャンプ
//[ 引  数 ]　PrevOrNext:-1=前方検索、+1:後方検索
//*****************************************************************************
procedure TKeyBinding.JumpNextMethod(PrevOrNext: Integer);
var
  i: Integer;
  Row, CommentRow: Integer;
  Index: Integer;
begin
  Row := EditView.Position.Row;
  Index := -1;

  if PrevOrNext < 0 then
  begin
    //前方の検索
    for i := Low(FStartLineArray) to High(FStartLineArray) do
    begin
      if Row > FStartLineArray[i].EndRow then
        Index := i;
    end;
  end
  else
  begin
    //後方の検索
    for i := High(FStartLineArray) downto Low(FStartLineArray) do
    begin
      if Row < FStartLineArray[i].StartRow then
        Index := i;
    end;
  end;

  EditView.Block.Reset;

  if Index = -1 then
  begin
    if PrevOrNext < 0 then
      //ファイルの先頭行へ
      EditView.Position.Move(1,1)
    else
      //ファイルの最終行へ
      EditView.Position.Move(EditView.Position.LastRow, 1);
  end
  else
  begin
    //クラス宣言の行
    if FStartLineArray[Index].LineType = ltMethod then
    begin
      //コメントの先頭を画面の先頭にする
      CommentRow := GetCommentStart(FStartLineArray[Index].StartRow, FAllText);
      EditView.SetTopLeft(CommentRow, 1);
      //ロジックのスタート行へカーソルを移動
      EditView.Position.Move(FStartLineArray[Index].LogicStart, 1);
    end
    else
    begin
      EditView.SetTopLeft(FStartLineArray[Index].StartRow - 1, 1);
      EditView.Position.Move(FStartLineArray[Index].StartRow,1);
    end;
  end;

  //再描画
  EditView.Paint;
end;

//*****************************************************************************
//[ 概  要 ]　次のメソッドへ移動
//*****************************************************************************
procedure TKeyBinding.GoNextMethod(const Context: IOTAKeyContext; KeyCode: TShortCut;	var BindingResult: TKeyBindingResult);
begin
	BindingResult := krHandled;
  if not CheckEditView() then Exit;

  FAllText := GetAllText();
  FStartLineArray := nil;
  if GetStartLineArray(FAllText, FStartLineArray) = 0 then Exit;

  JumpNextMethod(Integer(Context.Context));

  //キーを押し続けたときのリピートによる慣性を止める
  WaitForKeyMsgOver();
//  OutputDebugLog('GoNextMethod '+ WaitForKeyMsgOver.ToString);
end;

//*****************************************************************************
//[ 概  要 ]　KeyMessageがなくなるまで待機
//[ 戻り値 ]　True:キーがリピートされている
//*****************************************************************************
function TKeyBinding.WaitForKeyMsgOver(): Boolean;
var
  Msg: TMsg;
begin
  Result := False;
  while PeekMessage(Msg, 0, WM_KEYDOWN, WM_KEYDOWN, PM_REMOVE) do
  begin
    if Msg.lParam and $40000000 <> 0 then
      Result := True;
  end;
end;

//*****************************************************************************
//[ 概  要 ]　選択行を含むメソッド全体を選択
//*****************************************************************************
procedure TKeyBinding.SelectMethod(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
var
  AllText: string;
  LineNo: Integer;
  MethodLineInfo: TMethodLineInfo;
begin
	BindingResult := krHandled;
  if not CheckEditView() then Exit;

  AllText := GetAllText();
  LineNo := EditView.Position.Row;

  MethodLineInfo := GetMethodLineInfo(LineNo, AllText);
  if MethodLineInfo.StartRow = 0 then Exit;

  SelectRows(MethodLineInfo.CommentStart, MethodLineInfo.EndRow);

  //再描画
  EditView.Paint;
end;

//*****************************************************************************
//[ 概  要 ]　←キーの時、行頭で前の行の末尾へ移動
//*****************************************************************************
procedure TKeyBinding.LeftKeyPress(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
begin
  BindingResult := krHandled;
  if not CheckEditView() then Exit;

  //行頭なら前行の行末へ
	if EditView.Position.Column = 1 then
  begin
    if EditView.Position.Row > 1 then
    begin
      EditView.Position.MoveRelative(-1, 0);
      EditView.Position.MoveEOL;
    end;
	end
	else
  begin
    //カーソル位置≦行末
		if EditView.Position.Column <= GetEOLColumn() then
    begin
      EditView.Position.MoveRelative(0, -1);
    end
    else
    begin
      EditView.Position.MoveEOL;
		end;
	end;

  //キーを押し続けたときのリピートによる慣性を止める
  WaitForKeyMsgOver();
//  OutputDebugLog('LeftKeyPress '+ WaitForKeyMsgOver.ToString);
end;

//*****************************************************************************
//[ 概  要 ]　→キーの時、行末で次の行の先頭へ移動
//*****************************************************************************
procedure TKeyBinding.RightKeyPress(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
begin
	BindingResult := krHandled;
  if not CheckEditView() then Exit;

  //カーソル位置＜行末
  if EditView.Position.Column < GetEOLColumn() - 1 then
  begin
    EditView.Position.MoveRelative(0, 1);
  end
  else
  begin
    if EditView.Position.Row < EditView.Position.LastRow then
    begin
      //次の行の先頭へ
      EditView.Position.MoveRelative(1, 0);
      EditView.Position.MoveBOL;
    end
    else
    begin
      EditView.Position.MoveEOL;
    end;
  end;

  //キーを押し続けたときのリピートによる慣性を止める
  WaitForKeyMsgOver();
//  OutputDebugLog('RightKeyPress '+ WaitForKeyMsgOver.ToString);
end;

//*****************************************************************************
//[ 概  要 ]　アドインタイプ
//*****************************************************************************
function TKeyBinding.GetBindingType: TBindingType;
begin
	Result := btPartial;
end;

//*****************************************************************************
//[ 概  要 ]　表示名
//*****************************************************************************
function TKeyBinding.GetDisplayName: string;
begin
	//パッケージのインストールやエディタオプションでの表示名。
	Result := 'ユーザ定義ShortCut';
end;

//*****************************************************************************
//[ 概  要 ]　識別名
//*****************************************************************************
function TKeyBinding.GetName: string;
begin
	//レジストリに登録されるモジュール名
	Result := 'IDEShortCutEx';
end;

//*****************************************************************************
//登録
//*****************************************************************************
procedure Register;
begin
	BindNo := (BorlandIDEServices as IOTAKeyBoardServices).AddKeyboardBinding(TKeyBinding.Create);
end;

//*****************************************************************************
//削除
//*****************************************************************************
procedure UnRegister;
begin
	(BorlandIDEServices as IOTAKeyBoardServices).RemoveKeyboardBinding(BindNo);
end;

//*****************************************************************************
//initialization/finalization
//*****************************************************************************
initialization
finalization
	UnRegister;
end.
