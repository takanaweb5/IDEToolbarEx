unit MethodForm;

interface

uses
  Windows, Messages, Controls, Classes, ToolsAPI, SysUtils, StdCtrls,
  Graphics, Forms, ExtCtrls, AppEvnts, ImgList, Menus, System.ImageList;

type
  TLineInfo = record
    LineNo: Integer;
    Command: string;
    Define: string;
    ClassName: string;
    MethodName: string;
    BookMark: set of (bsOn, bsGray);
    Color: TColor;
    Select: Boolean;
  end;

{ TSizeGrip }
  TSizeGrip = class(TPaintBox)
    procedure Paint(); override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
  private
    FOrgWndProc: TWndMethod;
    procedure SubClassProc(var Message: TMessage);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

{ TMethodListForm }
  TMethodListForm = class(TForm)
    ApplicationEvents: TApplicationEvents;
    ImageList: TImageList;
    ListBox: TListBox;
    cmdOK: TButton;
    chkMethodName: TCheckBox;
    chkSort: TCheckBox;
    chkClassName: TCheckBox;
    cmbCommand: TComboBox;
    PopupMenu: TPopupMenu;
    miBookMarkSet: TMenuItem;
    miBookMarkDel: TMenuItem;
    miGoBack: TMenuItem;
    N1: TMenuItem;
//    procedure FormResize(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure ApplicationEventsIdle(Sender: TObject; var Done: Boolean);
    procedure miBookMarkSetClick(Sender: TObject);
    procedure miBookMarkDelClick(Sender: TObject);
    procedure miGoBackClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ListBoxDrawItem(Control: TWinControl; Index: Integer; ARect: TRect; State: TOwnerDrawState);
    procedure ListBoxClick(Sender: TObject);
    procedure ListBoxDblClick(Sender: TObject);
    procedure ListBoxMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure ListBoxKeyPress(Sender: TObject; var Key: Char);
    procedure cmdOKClick(Sender: TObject);
  private
//    SizeGrip: TSizeGrip;
    FimplementationRow: Integer;//implementationのある行
    FinterfaceRow: Integer; //interfaceの開始行
    FInfoIndex: Integer;    //FMethodListの起動時のカーソル位置のインデックス
    FTopRow: Integer;       //フォーム起動時の画面の先頭行
    FEditPos: TOTAEditPos;  //フォーム起動時のカーソル位置
    FText: string;          //編集中のファイルのすべての行
    FMethodList: array of TLineInfo;
//    procedure WMMove(var Message: TWMMove); message WM_MOVE;
    procedure ListBoxClickBookMark(Index: Integer);
    procedure ListBoxClickList(Index: Integer);
    function SetMethodList(const Text: string): Integer;
    function SetMethodLines(const Text: string): Integer;
    procedure SetBookMarkInfo();
    procedure SetLstBox(Sender: TObject);
    procedure ClearColor();
    function GetLineInfo(Index: Integer): TLineInfo;
    procedure SetColor(ColorToggle: Boolean);
    procedure ClearBookmark(Index: Integer);
    procedure JumpSelectMethod(Index: Integer);
  public
  end;

var
  MethodListForm: TMethodListForm;

implementation
{$R *.dfm}

uses
  ComObj, StrUtils, Dialogs, GlobalUnit, TypInfo;

type
{ TListBoxEx }
//  TBookMarkClickEvent = procedure(Sender: TObject; Index: Integer) of object;
  TListBoxEx = class(StdCtrls.TListBox)
  private
    procedure CNDrawItem(var Message: TWMDrawItem); message CN_DRAWITEM;
    procedure WMMouseWheel(var Message: TWMMouseWheel); message WM_MOUSEWHEEL;
  public
    constructor CreateClone(Original: TListBox);
  end;

const
  LEFTWIDTH = 18;
  C_CLASSDEFINE = 'ClassDefine';
  DrawingColor: array[Boolean] of TColor = ($FFFFCC, $CCFFFF);

{ TSizeGrip }
//*****************************************************************************
//[ 概  要 ]　サイズグリップを作成する
//*****************************************************************************
constructor TSizeGrip.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Parent  := AOwner as TWinControl;
  Cursor  := crSizeNWSE;
  Width   := GetSystemMetrics(SM_CXVSCROLL);
  Height  := GetSystemMetrics(SM_CXHSCROLL);
  Left    := Parent.ClientWidth  - Width;
  Top     := Parent.ClientHeight - Height;
  Anchors := [akRight, akBottom];

  //サブクラス化
  FOrgWndProc := Parent.WindowProc;
  Parent.WindowProc := SubClassProc;
end;

//*****************************************************************************
//[ 概  要 ]　デストラクタ
//*****************************************************************************
destructor TSizeGrip.Destroy;
begin
  if (Parent <> nil) and Parent.HandleAllocated then
    Parent.WindowProc := FOrgWndProc;
  inherited Destroy;
end;

//*****************************************************************************
//[ 概  要 ]　フォームのサブクラス化
// フォーム移動時、サイズグリップを確実に再描画させる　
// フォームの最大化時、サイズグリップを表示させない　
//*****************************************************************************
procedure TSizeGrip.SubClassProc(var Message: TMessage);
begin
  case Message.Msg of
    WM_MOVE: Invalidate;
    WM_SIZE: Visible := not IsZoomed(Parent.Handle);
  end;
  FOrgWndProc(Message);
end;

//*****************************************************************************
//[イベント]　サイズグリップの描画処理
//*****************************************************************************
procedure TSizeGrip.Paint();
begin
  with Self.Canvas do
  begin
    DrawFrameControl(Handle, ClipRect, DFC_SCROLL, DFCS_SCROLLSIZEGRIP);
  end;
end;

//*****************************************************************************
//[イベント]　サイズグリップのMouseDown時
//[ 概  要 ]　サイズの変更
//*****************************************************************************
procedure TSizeGrip.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if ssLeft in Shift then
  begin
    ReleaseCapture();
    Parent.Perform(WM_SYSCOMMAND, SC_SIZE or WMSZ_BOTTOMRIGHT,0);
  end;
  inherited;
end;


{ TListBoxEx }
//*****************************************************************************
//[コンストラクタ]　
//[ 概  要 ]　TListBoxをTListBoxExに変身させる
//*****************************************************************************
constructor TListBoxEx.CreateClone(Original: TListBox);
  procedure CopyEvents();
  var
    i, Count:  Integer;
    PropList:  PPropList;
    PropInfo:  TPropInfo;
    PropValue: TMethod;
  begin
    PropList := nil;
    Count := GetPropList(Original, PropList);
    try
      //プロパティの数だけループ
      for i := 0 to Count - 1 do
      begin
        PropInfo := PropList[i]^;
        //プロパティがイベントの時
        if PropInfo.PropType^.Kind = tkMethod  then
        begin
          PropValue := GetMethodProp(Original, PropInfo.Name);
          SetMethodProp(Self, PropInfo.Name, PropValue);
        end;
      end;
    finally
      if PropList <> nil then FreeMem(PropList);
    end;
  end;
var
  MemStream: TMemoryStream;
begin
  inherited Create(Original.Owner);
  Self.Parent := Original.Parent;
  MemStream := TMemoryStream.Create;
  try
    MemStream.WriteComponent(Original);
    Original.Name := ''; //これがないと重複例外が発生
    MemStream.Position := 0;
    MemStream.ReadComponent(Self);  //イベントはコピーされない

    //イベントのコピー
    CopyEvents();

//    Self.OnDrawItem  := Original.OnDrawItem;
//    Self.OnMouseDown := Original.OnMouseDown;
//    Self.OnClick     := Original.OnClick;
//    Self.OnDblClick  := Original.OnDblClick;
//    Self.OnKeyDown   := Original.OnKeyDown;
//    Self.OnKeyPress  := Original.OnKeyPress;

    Original.Free;
  finally
    MemStream.Free;
  end;
end;

//*****************************************************************************
//[ 概  要 ] リストボックスの選択行のフォーカスの四角を描画させない　
//*****************************************************************************
procedure TListBoxEx.CNDrawItem(var Message: TWMDrawItem);
var
  State: TOwnerDrawState;
begin
  with Message.DrawItemStruct^ do
  begin
    State := TOwnerDrawState(LongRec(itemState).Lo);
    Canvas.Handle := hDC;
    Canvas.Font := Font;
    Canvas.Brush := Brush;
//    if (Integer(itemID) >= 0) and (odSelected in State) then
//    begin
//      Canvas.Brush.Color := clHighlight;
//      Canvas.Font.Color := clHighlightText
//    end;
    if Integer(itemID) >= 0 then
      DrawItem(itemID, rcItem, State) else
      Canvas.FillRect(rcItem);
//    if odFocused in State then DrawFocusRect(hDC, rcItem);
    Canvas.Handle := 0;
  end;
end;

//*****************************************************************************
//[イベント]　リストボックスマウスホイール時
//[ 概  要 ]　リストボックスを気持ちよくスクロールさせる
//*****************************************************************************
procedure TListBoxEx.WMMouseWheel(var Message: TWMMouseWheel);
begin
  if Message.WheelDelta < 0 then
  begin
    Self.TopIndex := Self.TopIndex + 3;
  end
  else
  begin
    if Self.TopIndex < 4 then
      Self.TopIndex := 0
    else
      Self.TopIndex := Self.TopIndex - 3;
  end;
end;

//*****************************************************************************
//[ 概  要 ]　ソート用のユーザ定義関数
//*****************************************************************************
function StringListCompareStrings(List: TStringList; Index1, Index2: Integer): Integer;
var
  Info1, Info2: TLineInfo;
begin
  Info1 := MethodListForm.FMethodList[Integer(List.Objects[Index1])];
  Info2 := MethodListForm.FMethodList[Integer(List.Objects[Index2])];

  //クラス名でソート
  Result := AnsiCompareText(Info1.ClassName, Info2.ClassName);
  if Result <> 0 then Exit;

  //メソッド名でソート
  Result := AnsiCompareText(Info1.MethodName, Info2.MethodName);
  if Result <> 0 then Exit;

  //行番号でソート
  Result := Info1.LineNo - Info2.LineNo;
end;

{ TMethodListForm }
//*****************************************************************************
//[イベント]　アプリケーションがアイドルの時
//[ 概  要 ]　各コントロールのEnabledを設定する
//*****************************************************************************
procedure TMethodListForm.ApplicationEventsIdle(Sender: TObject; var Done: Boolean);
begin
try
  //当フォーム以外のApplicationEventsIdleを実行させない
  ApplicationEvents.CancelDispatch;

  chkSort.Enabled := (ListBox.Items.Count > 1);

  if ListBox.Items.Count = 0 then
  begin
    cmdOK.Enabled         := False;
    chkClassName.Enabled  := chkClassName.Checked;
    chkMethodName.Enabled := chkMethodName.Checked;
    miBookMarkSet.Enabled := False;
    miBookMarkDel.Enabled := False;
  end
  else
  begin
    if ListBox.ItemIndex = -1 then
    begin
      cmdOK.Enabled         := False;
      chkClassName.Enabled  := chkClassName.Checked;
      chkMethodName.Enabled := chkMethodName.Checked;
      miBookMarkSet.Enabled := False;
      miBookMarkDel.Enabled := False;
    end
    else
    begin
      cmdOK.Enabled         := True;
      chkClassName.Enabled  := True;
      chkMethodName.Enabled := True;
      miBookMarkSet.Enabled := not(bsOn in GetLineInfo(ListBox.ItemIndex).BookMark);
      miBookMarkDel.Enabled := not(GetLineInfo(ListBox.ItemIndex).BookMark = []);
    end;
  end;
except end;
end;

//*****************************************************************************
//[イベント]　フォーム作成時
//[ 概  要 ]　初期処理
//*****************************************************************************
procedure TMethodListForm.FormCreate(Sender: TObject);
begin
  //サイズグリップの作成
  TSizeGrip.Create(Self);

  //リストボックスを変身させる
  ListBox := TListBoxEx.CreateClone(ListBox);
  ListBox.ItemHeight := ListBox.ItemHeight + 4;

  //フォームの最小幅の設定
  Constraints.MinWidth := 420;
end;

//*****************************************************************************
//[イベント]　フォーム表示時
//[ 概  要 ]　画面の初期設定
//*****************************************************************************
procedure TMethodListForm.FormShow(Sender: TObject);
  function GetLineIndex(LineNo: Integer): Integer;
  var
    i: Integer;
  begin
    Result := -1;

    //LineNoを含むメソッドを検索
    for i := ListBox.Items.Count -1  downto  0 do
    begin
      if GetLineInfo(i).LineNo <= LineNo then
      begin
        Result := i;
        Exit;
      end;
    end;
  end;
var
  LineNo: Integer;
begin
  ListBox.Clear;

  //ApplicationEventsIdleを有効にする
  ApplicationEvents.OnIdle := ApplicationEventsIdle;
  ApplicationEvents.Activate;

  //起動時のカーソル位置情報を保存
  FTopRow  := EditView.TopRow;
  FEditPos := EditView.CursorPos;

  //編集中のファイル名をフォームのタイトルバーに表示
  Caption := 'メソッド一覧 - ' + ExtractFileName(EditView.Buffer.FileName);

  FText := GetAllText();

  //各コントロールのEnabledの設定
  if SetMethodList(FText) = 0 then
  begin
    cmbCommand.Enabled    := False;
    cmdOK.Enabled         := False;
    chkSort.Enabled       := False;
    chkClassName.Enabled  := False;
    chkMethodName.Enabled := False;
    Exit;
  end
  else
  begin
    cmbCommand.Enabled  := True;
    chkSort.Enabled     := True;
  end;

  chkSort.OnClick       := nil;
  chkClassName.OnClick  := nil;
  chkMethodName.OnClick := nil;
  cmbCommand.OnClick    := nil;

  chkSort.Checked       := False;
  chkClassName.Checked  := False;
  chkMethodName.Checked := False;
  cmbCommand.ItemIndex  := 0;

  SetLstBox(nil);

  chkSort.OnClick       := SetLstBox;
  chkClassName.OnClick  := SetLstBox;
  chkMethodName.OnClick := SetLstBox;
  cmbCommand.OnClick    := SetLstBox;

  //カーソル位置のメソッドを選択
  LineNo := EditView.Position.Row;
  if GetLineIndex(LineNo) <> -1 then
  begin
    ListBox.ItemIndex := GetLineIndex(LineNo);
    FInfoIndex := Integer(ListBox.Items.Objects[ListBox.ItemIndex]);
    FMethodList[FInfoIndex].Select := True;
  end
  else
  begin
    //先頭を選択
    ListBox.ItemIndex := 0;
    FInfoIndex := -1;
  end;
end;

//*****************************************************************************
//[イベント]　フォームをCloseする時
//*****************************************************************************
procedure TMethodListForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  case ModalResult of
  mrOk : ;//何もしない
//   begin
//    //選択されたメソッドのロジック開始位置へカーソルを移動
//    LineNo := GetLineInfo(ListBox.ItemIndex).LineNo;
//    LineNo := GetMethodLineInfo(LineNo, FText).LogicStart;
//    EditView.Position.Move(LineNo, 1);
//   end;
  mrCancel :
   begin
    //フォーム起動時のカーソル位置を表示
    EditView.SetTopLeft(FTopRow, 1);
    EditView.SetCursorPos(FEditPos);
   end;
  end;

  EditView.Paint;

  //ウィンドウのサイズを元に戻す
//  ShowWindow(Handle, SW_HIDE);
  Self.WindowState := wsNormal;

  FMethodList := nil;

  //ApplicationEventsIdleを無効にする
  ApplicationEvents.OnIdle := nil;
end;

//*****************************************************************************
//[イベント]　フォームをCloseする前に確認する
//*****************************************************************************
procedure TMethodListForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if ModalResult = mrOk then Exit;

  if (FTopRow <> EditView.TopRow) or (FEditPos.Line <> EditView.CursorPos.Line) then
  begin
    case MessageDlg('元の位置に戻しますか？', mtInformation,
                                      [mbYes, mbNo, mbCancel], 0) of
    mrYes   : ModalResult := mrCancel;
    mrNo    : ModalResult := mrOk;
    mrCancel: CanClose := False;
    end;
  end;
end;

//*****************************************************************************
//[イベント]　KeyDown時
//[ 概  要 ] ESCでフォームを閉じる　
//*****************************************************************************
procedure TMethodListForm.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_ESCAPE then ModalResult := mrCancel;
end;

//*****************************************************************************
//[イベント]　閉じるボタンクリック時
//*****************************************************************************
procedure TMethodListForm.cmdOKClick(Sender: TObject);
begin
  ModalResult := mrOk;
end;

//*****************************************************************************
//[イベント]　リストボックスKeyPress時
//[ 概  要 ]　①スペースキーでブックマークの設定/解除を行う
//            　※KeyDownではデフォルトの動作を抑制できないのでKeyPressを使用
//            ②アルファベットのキーで、該当するメソッドを検索
//*****************************************************************************
procedure TMethodListForm.ListBoxKeyPress(Sender: TObject; var Key: Char);
  function SerchMethod(ClassName: string; StartIndex: Integer): Integer;
  var
    i,k: Integer;
  begin
    Result := -1;
    for i := StartIndex to ListBox.Count - 1 do
    begin
      k := Integer(ListBox.Items.Objects[i]);
      if AnsiSameText(FMethodList[k].ClassName, ClassName) then
      begin
        //メソッドの1文字目を照合(MethodName=''の場合も例外なし)
        if AnsiSameText(Key, LeftStr(FMethodList[k].MethodName, 1)) then
        begin
          Result := i;
          Exit;
        end;
      end;
    end;
  end;
var
  Index: Integer;
  j, k: Integer;
  blnSerch: Boolean;
begin
  //リストボックスが選択されていないと対象外
  Index := ListBox.ItemIndex;
  if Index < 0 then Exit;

  //スペースキーでブックマークの設定/解除を行う
  if Key = ' ' then
  begin
    ListBoxClickBookMark(Index);
    //デフォルトの操作を抑制する
    Key := Chr(0);
    Exit;
  end;

  //アルファベットのキーで、該当するメソッドを検索
  case Key of
   'A'..'Z','a'..'z','0'..'9','_':
    begin
      k := Integer(ListBox.Items.Objects[Index]);

      blnSerch := (FMethodList[k].Command <> C_CLASSDEFINE);
      if blnSerch = False then
        if (cmbCommand.Text = 'すべて') and (chkSort.Checked = True) then
          blnSerch := True;
        if chkClassName.Checked then
          blnSerch := True;

      if blnSerch then
      begin
        //該当するクラスのKeyで始まるメソッドを選択されたIndex以降で検索
        j := SerchMethod(FMethodList[k].ClassName, Index + 1);
        if j < 0 then
          //選択されたIndex以降に見つからなければ最初から検索
          j := SerchMethod(FMethodList[k].ClassName, 0);

        if j >= 0 then
        begin
          ListBox.ItemIndex := j;
          JumpSelectMethod(j);
        end;
      end;

      //デフォルトの操作を抑制する
      Key := Chr(0);
    end;
  end;
end;

//*****************************************************************************
//[イベント] リストボックスKeyDown時
//[ 概  要 ] DELETEでブックマークを削除　
//*****************************************************************************
//procedure TMethodListForm.ListBoxKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
//begin
//  if Key = VK_DELETE then
//  begin
//  end;
//end;

//*****************************************************************************
//[ 概  要 ] リストボックスで選択されたメソッド内のブックマークをすべて削除する
//[ 引  数 ] リストボックスで選択されているIndex
//*****************************************************************************
procedure TMethodListForm.ClearBookmark(Index: Integer);
var
  i, k: Integer;
  BMLineNo: Integer;
begin
  k := Integer(ListBox.Items.Objects[Index]);
  if FMethodList[k].BookMark = [] then Exit;
  FMethodList[k].BookMark := [];

  //ブックマーク数ループ
  for i := 0 to 9 do
  begin
    BMLineNo := EditView.BookmarkPos[i].Line;
    if BMLineNo = 0 then Continue;
    if GetMethodLineInfo(BMLineNo, FText).StartRow = FMethodList[k].LineNo then
    begin
      //Bookmarkの解除
      EditView.BookmarkGoto(i);
      EditView.BookmarkToggle(i);
    end;
  end;

  //再描画
  EditView.Paint;
  ListBox.Repaint;
end;


//*****************************************************************************
//[ 概  要 ] リストボックスの項目の描画を独自に行う　
//*****************************************************************************
procedure TMethodListForm.ListBoxDrawItem(Control: TWinControl; Index: Integer; ARect: TRect; State: TOwnerDrawState);
  procedure DrawDotLine(Canvas: TCanvas; StartX, EndX: Integer);
  var
    X: Integer;
  begin
    X := StartX;
    while X < EndX do
    begin
      Canvas.Pixels[X, ARect.Top] := clGray;
      Inc(X, 2);
    end;
  end;
var
  R: TRect;
begin
  with ListBox do
  begin
    //*************************************
    //ブックマークアイコンエリアの描画
    //*************************************
    if Index = Items.Count - 1 then
      //最後の行の時は、残りの領域も対象
      R := Rect(0, ARect.Top, LEFTWIDTH, ClientRect.Bottom)
    else
      R := Rect(0, ARect.Top, LEFTWIDTH, ARect.Bottom);
    Canvas.Brush.Color := RGB(244, 244, 244);
    Canvas.FillRect(R);
    Canvas.Pen.Color := RGB(153, 153, 204);
    Canvas.MoveTo(LEFTWIDTH, ARect.Top);
    Canvas.LineTo(LEFTWIDTH, R.Bottom);

    //カレント行は、!マークをつける
    if GetLineInfo(Index).Select then
    begin
      Canvas.Font.Color := clRed;
      Canvas.TextOut(0, R.Top + 1, '!');
    end;

    //ブックマークのアイコンをつける
    if bsOn in GetLineInfo(Index).BookMark then
      ImageList.Draw(Canvas,3, ARect.Top, 0)
    else if bsGray in GetLineInfo(Index).BookMark then
      ImageList.Draw(Canvas,3, ARect.Top, 1);

    //*************************************
    //上記以外のエリアの描画
    //*************************************
    R := Rect(ARect.Left + LEFTWIDTH + 1, ARect.Top, ARect.Right, ARect.Bottom);

    Canvas.Brush.Color := GetLineInfo(Index).Color;

    if cmbCommand.Text = 'クラス宣言' then
        Canvas.Brush.Color := $FFDDFE
    else
      if (chkSort.Checked = False) and (chkClassName.Checked = False) and
         (GetLineInfo(Index).Command = C_CLASSDEFINE) then
        Canvas.Brush.Color := $FFDDFE;

    //フォーカスのある行の色を変更する
    if odSelected in State then
      Canvas.Brush.Color := Canvas.Brush.Color - $0C0C0C;

    Canvas.FillRect(R);
    Canvas.Font.Color := clBlack;
    Canvas.TextOut(R.Left + 6, R.Top + 1, Items[Index]);

    //フォーカスの点線枠
    if odSelected in State then Canvas.DrawFocusRect(R);

    //implementation部の開始前に下線を引く
    if (chkSort.Checked = False) and (chkClassName.Checked = False) then
    begin
      if (Index > 0) and
         (GetLineInfo(Index - 1).LineNo < FimplementationRow) and
         (FimplementationRow < GetLineInfo(Index).LineNo) then
      begin
        //極細の点線を描画
        DrawDotLine(Canvas, LEFTWIDTH + 1, ARect.Right);
//        Canvas.Pen.Color := clBlack;
//        Canvas.MoveTo(LEFTWIDTH + 1, ARect.Top);
//        Canvas.LineTo(ARect.Right, ARect.Top);
      end;
    end;
  end;
end;

//*****************************************************************************
//[イベント]　リストボックスクリック時
//[ 概  要 ]　選択されたメソッドを表示する
//*****************************************************************************
procedure TMethodListForm.ListBoxClick(Sender: TObject);
begin
  JumpSelectMethod(ListBox.ItemIndex);
end;

//*****************************************************************************
//[ 概  要 ]　リストボックの選択されたメソッドを表示する
//[ 引  数 ]　ListBoxIndex
//*****************************************************************************
procedure TMethodListForm.JumpSelectMethod(Index: Integer);
var
  LineNo: Integer;
  MethodLineInfo: TMethodLineInfo;
begin
  LineNo := GetLineInfo(Index).LineNo;

  if GetLineInfo(Index).Command = C_CLASSDEFINE then
  begin
    //再描画
    EditView.SetTopLeft(GetCommentStart(LineNo, FText), 1);
    EditView.Paint;
    EditView.Position.Move(LineNo, 1);
  end
  else
  begin
    MethodLineInfo := GetMethodLineInfo(LineNo, FText);
    if MethodLineInfo.StartRow = 0 then Exit;

    //再描画
    EditView.SetTopLeft(MethodLineInfo.CommentStart, 1);
    EditView.Paint;
    EditView.Position.Move(MethodLineInfo.LogicStart, 1);
  end;
end;

//*****************************************************************************
//[イベント]　リストボックスダブルクリック時
//*****************************************************************************
procedure TMethodListForm.ListBoxDblClick(Sender: TObject);
begin
  ModalResult := mrOk;
end;

//*****************************************************************************
//[イベント]　リストボックスマウスダウン時
//[ 概  要 ]　クリックされた領域で制御を変える
//*****************************************************************************
procedure TMethodListForm.ListBoxMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  Index: Integer;
begin
  if Button = mbLeft then
  begin
    Index := ListBox.ItemAtPos(Point(X,Y),True);
    if Index <> -1 then
    begin
      if X < LEFTWIDTH then
      begin
        ListBoxClickBookMark(Index);
      end
      else
      begin
        ListBoxClickList(Index);
      end
    end
  end;
end;

//*****************************************************************************
//[イベント]　ブックマークアイコンエリアのクリック時
//[ 概  要 ]　ブックマークアイコンの設定
//*****************************************************************************
procedure TMethodListForm.ListBoxClickBookMark(Index: Integer);
var
  i, k: Integer;
  BMLineNo: Integer;
  BookMarkID: Integer;
  MethodLineInfo: TMethodLineInfo;
begin
  k := Integer(ListBox.Items.Objects[Index]);

  if bsOn in GetLineInfo(Index).BookMark then
    //すでに設定されていれば解除
    Exclude(FMethodList[k].BookMark, bsOn)
  else
  begin
    //未設定であれば設定
    if GetUnusedBookMarkID() >= 10 then
    begin
      MessageDlg('しおりは10個までしか設定できません', mtInformation, [mbOK], 0);
      Exit
    end;
    Include(FMethodList[k].BookMark, bsOn);
  end;

  if GetLineInfo(Index).Command = C_CLASSDEFINE then
    with MethodLineInfo do
    begin
      StartRow   := FMethodList[k].LineNo;
      LogicStart := StartRow;
    end
  else
    MethodLineInfo := GetMethodLineInfo(FMethodList[k].LineNo, FText);

  if bsOn in GetLineInfo(Index).BookMark then
  begin
    //しおりの設定
    BookMarkID := GetUnusedBookMarkID();
    EditView.Position.Move(MethodLineInfo.StartRow, 1);
    EditView.BookmarkRecord(BookMarkID);
  end
  else
  begin
    //ブックマーク数ループ
    for i := 0 to 9 do
    begin
      BMLineNo := EditView.BookmarkPos[i].Line;
      if BMLineNo = 0 then Continue;
      if (BMLineNo = MethodLineInfo.StartRow) or
         (BMLineNo = MethodLineInfo.LogicStart) then
      begin
        //Bookmarkの解除
        EditView.BookmarkGoto(i);
        EditView.BookmarkToggle(i);
      end;
    end;
  end;

  //再描画
  EditView.Paint;
  ListBox.Repaint;

  //ダブルクリックイベントを抑止
  ListBox.OnDblClick := nil;
end;

//*****************************************************************************
//[イベント]　ブックマークアイコンエリア以外のクリック時
//[ 概  要 ]　ダブルクリックイベントを復活
//*****************************************************************************
procedure TMethodListForm.ListBoxClickList(Index: Integer);
begin
  ListBox.OnDblClick := ListBoxDblClick;
end;

//*****************************************************************************
//[イベント]  フォーム表示時
//          　ソートクリック時
//          　チェックボックスチェック時
//[ 概  要 ]　FMethodListの内容からリストボックスを設定する
//*****************************************************************************
procedure TMethodListForm.SetLstBox(Sender: TObject);
var
  i: Integer;
  str: string;
  ST: TStringList;
  strClassName, strMethodName: string;
  iLineNo: Integer;
  iColor: TColor;
begin
  iColor := clWhite;

  //選択されている行の情報を取得する
  iLineNo := 0;
  if ListBox.ItemIndex >= 0 then
  begin
    with GetLineInfo(ListBox.ItemIndex) do
    begin
      iLineNo       := LineNo;
      strClassName  := ClassName;
      strMethodName := MethodName;
      iColor        := Color;
    end;
  end;

  //対象の行を選択する
  ST := TStringList.Create;
  try
    for i := 0 to Length(FMethodList) - 1 do
    begin
      if FMethodList[i].Command = C_CLASSDEFINE then
      begin
        if chkClassName.Checked then
          if AnsiSameText(FMethodList[i].ClassName,  strClassName) = False then
            Continue;
        if chkMethodName.Checked then Continue;
        if (cmbCommand.Text <> 'すべて') and (cmbCommand.Text <> 'クラス宣言')  then
          Continue;

        if cmbCommand.Text = 'クラス宣言' then
          str := Format('%-18s = %s', [FMethodList[i].ClassName, FMethodList[i].Define])
        else
          if chkClassName.Checked or chkSort.Checked then
//            str := Format('%-11s %s = %s', [' ', FMethodList[i].ClassName, FMethodList[i].Define])
            str := Format('%s  = %s', [FMethodList[i].ClassName, FMethodList[i].Define])
          else
            str := Format('%-18s = %s', [FMethodList[i].ClassName, FMethodList[i].Define]);
      end
      else
      begin
        if chkClassName.Checked then
          if AnsiSameText(FMethodList[i].ClassName,  strClassName) = False then
            Continue;
        if chkMethodName.Checked then
          if AnsiSameText(FMethodList[i].MethodName,strMethodName) = False then
            Continue;
        if cmbCommand.ItemIndex > 0 then
          if AnsiSameText(FMethodList[i].Command, cmbCommand.Text) = False then
            Continue;

        str := Format('%-11s %s', [FMethodList[i].Command, FMethodList[i].Define]);
      end;

      ST.AddObject(str, Pointer(i));
    end;

    //ソートがチェックされていればソートさせる
    if chkSort.Checked then ST.CustomSort(StringListCompareStrings);

    ListBox.Items.Assign(ST);
  finally
    ST.Free;
  end;

  //選択されていた行を再選択させる
  if iLineNo <> 0 then
  begin
    for i := 0 to ListBox.Items.Count - 1 do
    begin
      if GetLineInfo(i).LineNo = iLineNo then
      begin
        ListBox.ItemIndex := i;
        Break;
      end;
    end;
  end;

  //背景色の情報を設定する
  ClearColor();
  SetColor(True);

  //選択行の背景色が選択前と変わったら選択前と同じになるようにやり直し
  if ListBox.ItemIndex >= 0 then
  begin
    if (iColor <> clWhite) and
       (iColor <> GetLineInfo(ListBox.ItemIndex).Color) then
    begin
      //背景色の情報を再設定する
      SetColor(False);
    end;
  end;

  //再描画
  ListBox.Repaint;
  ListBox.SetFocus;
end;

//*****************************************************************************
//[ 概  要 ]　背景色を設定する
//[ 引  数 ]　最初の色をいずれにするかのトグル
// クラス名が初期値の時は「白」
// クラス名が同じ間は、同一色で統一
//*****************************************************************************
procedure TMethodListForm.SetColor(ColorToggle: Boolean);
  function GetLastName(Index: Integer): string;
  var
    i: Integer;
  begin
    for i := Index - 1 downto 0 do
    begin
      if GetLineInfo(i).ClassName <> '' then
      begin
        Result := GetLineInfo(i).ClassName;
        Exit;
      end;
    end;
    Result := '';
  end;
var
  i, j: integer;
begin
  for i := 0 to ListBox.Items.Count - 1 do
  begin
    j := Integer(ListBox.Items.Objects[i]);
    if GetLineInfo(i).ClassName = '' then
    begin
      FMethodList[j].Color := clWhite;
      Continue;
    end;

    //直前のクラス名と違う時、色を変更
    if AnsiSameText(GetLastName(i), GetLineInfo(i).ClassName) = False then
      ColorToggle := not ColorToggle;

    FMethodList[j].Color := DrawingColor[ColorToggle];
  end;
end;

//*****************************************************************************
//[ 概  要 ]　FMethodListを設定する
//[ 引  数 ]　編集ファイルの全行
//[ 戻り値 ]　(procedure or function) の行数
//*****************************************************************************
function TMethodListForm.SetMethodList(const Text: string): Integer;
begin
  //LineNo/Command/Define/ClassName/MethodName を設定
  Result := SetMethodLines(Text);
  if Result = 0 then Exit;

  //FMethodListのColorをクリアする
  ClearColor();

  //FMethodListのBookMarkを設定する
  SetBookMarkInfo();
end;

//*****************************************************************************
//[ 概  要 ]　FMethodListの Color(背景色)をクリアする
//*****************************************************************************
procedure TMethodListForm.ClearColor();
var
  i: Integer;
begin
  for i := 0 to Length(FMethodList) - 1 do
  begin
    FMethodList[i].Color := clWhite;
  end;
end;

//*****************************************************************************
//[ 概  要 ]　FMethodListの LineNo/Command/Define/ClassName/MethodName を設定
//[ 引  数 ]　編集ファイルの全行
//[ 戻り値 ]　(procedure or function) の行数
//*****************************************************************************
function TMethodListForm.SetMethodLines(const Text: string): Integer;
var
  i, j: Integer;
  ST: TStringList;
  TmpArray: array of Byte;
begin
try
  ST := TStringList.Create();
  try
    ST.Text := Text;

    //行数とIndexをあわせるため先頭に1行挿入
    ST.Insert(0, '');

    //行数分ループ
    for i := 1 to ST.Count do
    begin
      if LowerCase(LeftStr(Trim(ST[i]), 9)) = 'interface' then
        FinterfaceRow := i;
      if LowerCase(LeftStr(Trim(ST[i]),14)) = 'implementation' then
      begin
        FimplementationRow := i;
        Break;
      end
    end;

    SetLength(TmpArray, ST.Count);

    //該当行数の取得
    j := 0;
    //行数分ループ
    RegExp.Pattern := C_CLASS;
    for i := FinterfaceRow + 1 to ST.Count - 1 do
    begin
      if RegExp.Test(ST[i]) then
      begin
        TmpArray[i] := 1;
        Inc(j);
      end;
    end;

    //行数分ループ(implementation節)
    RegExp.Pattern := C_METHOD;
    for i := FimplementationRow + 1 to ST.Count - 1 do
    begin
      //forward宣言は対象外
      if LowerCase(RightStr(Trim(ST[i]),8)) = 'forward;' then Continue;
      if RegExp.Test(ST[i]) then
      begin
        TmpArray[i] := 2;
        Inc(j);
      end;
    end;

    Result := j;
    if  Result = 0 then Exit;

    SetLength(FMethodList, j);

    j := -1;
    //行数分ループし、FMethodList[j]を設定
    for i := 1 to ST.Count - 1 do
    begin
      case TmpArray[i] of
      1:
      begin
        RegExp.Pattern := C_CLASS;
        Match := RegExp.Execute(ST[i]).Item[0];
        Inc(j);
        with FMethodList[j] do
        begin
          LineNo  := i;
          Command := C_CLASSDEFINE; //
          ClassName := Match.SubMatches[0]; //例:TMethodListForm
          Define  := Match.SubMatches[1]; //例:class(TForm)
        end;
      end;
      2:
      begin
        RegExp.Pattern := C_METHOD;
        Match := RegExp.Execute(ST[i]).Item[0];
        Inc(j);
        with FMethodList[j] do
        begin
          LineNo  := i;
          Command := Match.SubMatches[0]; //例:procedure
          Define  := Match.SubMatches[1]; //例:TMethodListForm.SetMethodList(const Text: string);

          if Match.SubMatches[5] = '' then
          begin
            //例:TMethodListForm.SetMethodList(const Text: string);
            ClassName  := Match.SubMatches[3]; //例:TMethodListForm
            MethodName := Match.SubMatches[4]; //例:SetMethodList
          end
          else
          begin
            //例:StringListCompareStrings(Index1, Index2: Integer): Integer;
            ClassName  := '';
            MethodName := Match.SubMatches[5]; //例:StringListCompareStrings
          end;
        end;
      end;
    end;
  end;
  finally
    ST.Free;
  end;
except
  Result := 0;
end;
end;

//*****************************************************************************
//[ 概  要 ]　FMethodListの BookMark を設定
//*****************************************************************************
procedure TMethodListForm.SetBookMarkInfo();
var
  i, j: Integer;
  BMLineNo: Integer;
  MethodLineInfo: TMethodLineInfo;
begin
  //一旦すべてをクリア
  for i := 0 to Length(FMethodList) - 1 do
  begin
    FMethodList[i].BookMark := [];
  end;

  //ブックマーク数ループ
  for i := 0 to 9 do
  begin
    BMLineNo := EditView.BookmarkPos[i].Line;
    if BMLineNo <> 0 then
    begin
      MethodLineInfo := GetMethodLineInfo(BMLineNo, FText);
      if MethodLineInfo.StartRow <> 0 then
      begin
        //メソッド数分ループ
        for j := 0 to Length(FMethodList) - 1 do
        begin
          if MethodLineInfo.StartRow = FMethodList[j].LineNo then
          begin
            if (BMLineNo = MethodLineInfo.StartRow) or
               (BMLineNo = MethodLineInfo.LogicStart) then
            begin
              Include(FMethodList[j].BookMark, bsOn);
            end
            else
            begin
              Include(FMethodList[j].BookMark, bsGray);
            end;
            Break;
          end;
        end;
      end;
    end;
  end;
end;

//*****************************************************************************
//[ 概  要 ]　ListBoxのIndexに該当するTLineInfoを取得
//[ 引  数 ]　ListBoxのIndex
//[ 戻り値 ]　TLineInfo
//*****************************************************************************
function TMethodListForm.GetLineInfo(Index: Integer): TLineInfo;
var
  i: Integer;
begin
  i := Integer(ListBox.Items.Objects[Index]);
  Result := FMethodList[i];
end;

//*****************************************************************************
//[ 概  要 ]　「しおりの設定」メニュークリック時
//*****************************************************************************
procedure TMethodListForm.miBookMarkSetClick(Sender: TObject);
begin
  if ListBox.ItemIndex >= 0 then
  begin
    if not (bsOn in GetLineInfo(ListBox.ItemIndex).BookMark) then
      ListBoxClickBookMark(ListBox.ItemIndex);
  end;
end;

//*****************************************************************************
//[ 概  要 ]　「しおりの削除」メニュークリック時
//            ListBoxでDELETEキー押下時
//*****************************************************************************
procedure TMethodListForm.miBookMarkDelClick(Sender: TObject);
begin
  if ListBox.ItemIndex >= 0 then
    ClearBookmark(ListBox.ItemIndex);
end;

//*****************************************************************************
//[ 概  要 ]　メニュー「カーソル位置へ戻る」クリック時
//*****************************************************************************
procedure TMethodListForm.miGoBackClick(Sender: TObject);
var
  i: Integer;
begin
  //フォーム起動時のカーソル位置を表示
  EditView.SetTopLeft(FTopRow, 1);
  EditView.SetCursorPos(FEditPos);

  if FInfoIndex <> -1 then
  begin
    for i := 0 to ListBox.Items.Count - 1 do
    begin
      if FInfoIndex = Integer(ListBox.Items.Objects[i]) then
      begin
        ListBox.ItemIndex := i;
        Break;
      end;
    end;
  end
  else
  begin
    ListBox.ItemIndex := ListBox.Items.Count - 1;
  end;

  //再描画
  EditView.Paint;
end;

end.
