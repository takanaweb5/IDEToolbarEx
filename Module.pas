unit Module;

interface

uses
  Windows, Classes, Controls, ImgList, AppEvnts, ComCtrls, ToolsAPI, Dialogs, Forms, System.SysUtils,
  System.ImageList, System.Rtti, System.Actions, Vcl.ActnList, Vcl.ExtCtrls, Clipbrd, Vcl.Graphics, Vcl.Menus;

procedure Register;

type
  TMainModule = class(TDataModule)
    IconList: TImageList;
    ApplicationEvents: TApplicationEvents;
    ActionList: TActionList;
    procedure ApplicationEventsIdle(Sender: TObject; var Done: Boolean);
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
    procedure OnSetBookmarkClick(Sender: TObject);
    procedure OnNextBookmarkClick(Sender: TObject);
    procedure OnIndentClick(Sender: TObject);
    procedure OnCommentClick(Sender: TObject);
    procedure OnShowMethodFormClick(Sender: TObject);
  private
    FActions: array of TAction;
    FSearchText: string;
    procedure SetActionEnabled(ActionName: string; Enabled: Boolean);
  public
    procedure Indent(Sign: Integer);
  end;

var
  MainModule: TMainModule;

 implementation
{$R *.dfm}

uses
  MethodForm, GlobalUnit;

const
  TOOLBAR_NAME = 'IDEEditToolBarEx';

//*****************************************************************************
//[ 概  要 ]　Toolbarとその中のToolButtonの作成を行う
//*****************************************************************************
procedure TMainModule.DataModuleCreate(Sender: TObject);
const
  ButtonName: array[0..7] of string = (
    'SetBookmark'  ,
    'NextBookmark' ,
    'PrevBookmark' ,
    'Indent'       ,
    'Outdent'      ,
    'Comment'      ,
    'UnComment'    ,
    'ShowMethods');
  ButtonHint: array[0..7] of string = (
    'しおりの設定/解除'  ,
    '直後のしおり' ,
    '直前のしおり' ,
    'インデント'       ,
    'インデント解除'      ,
    'コメント化'      ,
    'コメント解除'    ,
    'メソッド一覧');
var
  IDEServices: INTAServices;
  ToolBar: TToolBar;
  ToolButton: TToolButton;
  i: Integer;
  Events: array[0..7] of TNotifyEvent;
  Icon: TIcon;
  function GetIndexFromActionList(ActionName: string): Integer;
  var
    i: Integer;
  begin
    Result := -1;
    for i := 0 to IDEServices.ActionList.ActionCount - 1 do
    begin
      if IDEServices.ActionList[i].Name = ActionName then
      begin
        Result := i;
        Exit;
      end;
    end;
  end;
begin
  FSearchText := '';
  IDEServices := (BorlandIDEServices as INTAServices);
  ToolBar := IDEServices.NewToolbar(TOOLBAR_NAME, 'ユーザ定義ツールバー');
  Assert(Assigned(ToolBar), 'ユーザ定義のツールバーの作成に失敗しました');
  ToolBar.Images := IconList;

  SetLength(FActions, 8);
  for i := 0 to High(FActions) do
  begin
    FActions[i] := TAction.Create(ActionList);
    FActions[i].ActionList := ActionList;
  end;
  Events[0] := OnSetBookmarkClick;
  Events[1] := OnNextBookmarkClick;
  Events[2] := OnNextBookmarkClick;
  Events[3] := OnIndentClick;
  Events[4] := OnIndentClick;
  Events[5] := OnCommentClick;
  Events[6] := OnCommentClick;
  Events[7] := OnShowMethodFormClick;

  //第4引数をTrueにすればStyle=tbsDividerとなるが、第3引数の設定が不要になる
  //第4引数をFalseにすればTSpeedButtonが作成されるみたいだ
  for i := 0 to High(FActions) do
  begin
    //セパレータの設定
//    case i of
//      3:begin
//        IDEServices.AddToolButton(TOOLBAR_NAME,'Divider1',nil ,True);
//        IDEServices.AddToolButton(TOOLBAR_NAME,'Divider2',nil ,True);
//      end;
//      5: IDEServices.AddToolButton(TOOLBAR_NAME,'Divider3',nil ,True);
//    end;
    ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME ,ButtonName[i] ,nil ,True) as TToolButton;
    with FActions[i] do
    begin
      Category   := 'ユーザ定義ツールバー';
      Name       := ButtonName[i];
      Hint       := ButtonHint[i];
      Caption    := ButtonHint[i];
      OnExecute  := Events[i];
      case i of
        7: ShortCut := TextToShortCut('F10');
      end;
    end;
//    //IDEServices.ActionListとのひもづけを行う
//    IDEServices.AddActionMenu('', FActions[i] , nil);
    Icon := TIcon.Create;
    try
      IconList.GetIcon(i,Icon);
      FActions[i].ImageIndex := i;
    finally
      Icon.Free;
    end;
    with ToolButton do
    begin
      Style      := tbsButton;
      Action     := FActions[i];
      case i of
        3,4: Visible := false;
      end;
    end;
  end;

//  //次を検索ボタン
//  i := GetIndexFromActionList('SearchAgainCommand');
//  if i <> -1 then
//  begin
//    ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME ,'FindNext' ,nil, True, 'Divider1') as TToolButton;
//    with ToolButton do
//    begin
//      Style      := tbsButton;
//      Action     := IDEServices.ActionList[i];
//    end;
//  end;
//  //検索ボタン
//  i := GetIndexFromActionList('SearchFindCommand');
//  if i <> -1 then
//  begin
//    ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME ,'Find' ,nil ,True, 'Divider1') as TToolButton;
//    with ToolButton do
//    begin
//      Style      := tbsButton;
//      Action     := IDEServices.ActionList[i];
//    end;
//  end;

//    ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME,'ElseMenu',nil ,True) as TToolButton;
//  with ToolButton do
//  begin
//    Style := tbsButton;
//    DropdownMenu := ElseMenu;
//    ImageIndex := 11;
//  end;

  //初期設定
  ToolBar.Perform(CM_RECREATEWND, 0, 0);//ボタンが変にならないようにするおまじない
  ToolBar.Visible := True;
  MethodListForm := TMethodListForm.Create(Self);
end;

//*****************************************************************************
//[ 概  要 ]　Toolbarを破棄する
//*****************************************************************************
procedure TMainModule.DataModuleDestroy(Sender: TObject);
var
  i: Integer;
begin
  MethodListForm.Free;
  for i := 0 to High(FActions) do
  begin
    FActions[i].Free;
  end;
end;

//*****************************************************************************
//[イベント]　しおりの設定/解除クリック時
//[ 概  要 ]　カーソル行のしおりの設定または解除を行う
//*****************************************************************************
procedure TMainModule.OnSetBookmarkClick(Sender: TObject);
var
  BookMarkID: Integer;
  Row: Integer;
begin
  if not CheckEditView() then Exit;

  //現在位置
  Row := EditView.Position.Row;

  BookMarkID := GetActLineBookMarkID();
  if BookMarkID >= 0 then
  begin
    //しおりの解除
    EditView.BookmarkToggle(BookMarkID);
  end
  else
  begin
    //しおりの設定
    BookMarkID := GetUnusedBookMarkID();
    if BookMarkID < 10 then
    begin
      //行の先頭にしおりを設定
      EditView.Position.Move(Row, 1);
      EditView.BookmarkRecord(BookMarkID);
     end
    else
    begin
      MessageDlg('しおりは10個までしか設定できません', mtInformation, [mbOK], 0);
    end;
  end;

  //再描画
  EditView.MoveViewToCursor;
  EditView.Paint;
end;

//*****************************************************************************
//[イベント]　次のしおり or 直前のしおりクリック時
//[ 概  要 ]　カーソル行直前または直後のしおりにジャンプ
//*****************************************************************************
procedure TMainModule.OnNextBookmarkClick(Sender: TObject);
var
  BookMarkID: Integer;
begin
  if not CheckEditView() then Exit;

  case (Sender as TAction).Name[1] of
  'N': //NextBookmark
    begin
      //直後のBookMarkIDを取得
      BookMarkID := GetNextBookmarkID(+1);
    end;
  'P': //PrevBookmark
    begin
      //直前のBookMarkIDを取得
      BookMarkID := GetNextBookmarkID(-1);
    end;
  else Exit;
  end;

  if BookMarkID >= 0 then
  begin
    EditView.BookmarkGoto(BookMarkID);
    EditView.MoveViewToCursor;

    //再描画
    EditView.Paint;
  end;
end;

//*****************************************************************************
//[イベント]　インデント or インデント解除クリック時
//[ 概  要 ]　選択行を字下げまたは字上げ
//*****************************************************************************
procedure TMainModule.OnIndentClick(Sender: TObject);
begin
  if not CheckEditView() then Exit;

  case (Sender as TAction).Name[1] of
  'I':Indent(+1);
  'O':Indent(-1);
  else Exit;
  end;

  //再描画
  EditView.Paint;
end;

//*****************************************************************************
//[イベント]　インデント or インデント解除
//[ 引  数 ]　+1:インデント、-1:インデント解除
//[ 概  要 ]　選択行を字下げまたは字上げ
//*****************************************************************************
procedure TMainModule.Indent(Sign: Integer);
var
  Indent: Integer;
  StartingRow, EndingRow: Integer;
  MinusRow: Integer;
begin
  //ブロック選択でカーソルが行頭ならその行は対象外にする
  if (EditView.Block.Style <> btLine) and
     (EditView.Block.Size <> 0) and
     (EditView.Block.EndingColumn = 1) then MinusRow := 1
                                       else MinusRow := 0;
  //選択を行選択に切り替え
  EditView.Block.Style := btLine;
  StartingRow  := EditView.Block.StartingRow;
  EndingRow    := EditView.Block.EndingRow - MinusRow;

  SelectRows(StartingRow, EndingRow);

  //インデントを実行
  Indent := EditView.Buffer.EditOptions.BlockIndent;
  EditView.Block.Indent(Indent * Sign);
end;

//*****************************************************************************
//[イベント]　コメント化 or コメント解除クリック時
//[ 概  要 ]　選択行をコメント化または解除
//*****************************************************************************
procedure TMainModule.OnCommentClick(Sender: TObject);
var
  i: Integer;
  StartingRow, EndingRow: Integer;
  MinusRow: Integer;
begin
  if not CheckEditView() then Exit;

  //ブロック選択でカーソルが行頭ならその行は対象外にする
  if (EditView.Block.Style <> btLine) and
     (EditView.Block.Size <> 0) and
     (EditView.Block.EndingColumn = 1) then MinusRow := 1
                                       else MinusRow := 0;
  //選択を行選択に切り替え
  EditView.Block.Style := btLine;
  StartingRow  := EditView.Block.StartingRow;
  EndingRow    := EditView.Block.EndingRow - MinusRow;

  for i := StartingRow to EndingRow do
  begin
    EditView.Position.Move(i, 1);
    case (Sender as TAction).Name[1] of
    'C': //Comment
      begin
        EditView.Position.InsertText('//');
      end;
    'U': //UnComment
      begin
        if EditView.Position.Read(2) = '//' then
            EditView.Position.Delete(2);
      end;
    end;
  end;

  //コメント行を選択
  SelectRows(StartingRow, EndingRow);

  //再描画
  EditView.Paint;
end;

//*****************************************************************************
//[イベント]　メソッド一覧クリック時
//[ 概  要 ]　メソッド一覧フォームを表示する
//*****************************************************************************
procedure TMainModule.OnShowMethodFormClick(Sender: TObject);
begin
  if not CheckEditView() then Exit;

  MethodListForm.ShowModal;
end;

//*****************************************************************************
//[イベント]　アプリケーションがアイドルの時
//[ 概  要 ]　各ToolButtonのEnabledを設定する
//*****************************************************************************
procedure TMainModule.ApplicationEventsIdle(Sender: TObject; var Done: Boolean);
var
  UsedBookMarkCount: Integer;
begin
  if not CheckEditView() then
  begin
    SetActionEnabled('SetBookmark' , False);
    SetActionEnabled('NextBookmark', False);
    SetActionEnabled('PrevBookmark', False);
    SetActionEnabled('Indent'      , False);
    SetActionEnabled('Outdent'     , False);
    SetActionEnabled('Comment'     , False);
    SetActionEnabled('UnComment'   , False);
    SetActionEnabled('ShowMethods' , False);
    Exit;
  end;
  UsedBookMarkCount := GetUsedBookMarkCount();
  SetActionEnabled('NextBookmark', (UsedBookMarkCount > 0));
  SetActionEnabled('PrevBookmark', (UsedBookMarkCount > 0));

  SetActionEnabled('SetBookmark', True);
  SetActionEnabled('Indent'     , True);
  SetActionEnabled('Outdent'    , True);
  SetActionEnabled('Comment'    , True);
  SetActionEnabled('UnComment'  , True);
  SetActionEnabled('ShowMethods', True);
end;

//*****************************************************************************
//[ 概  要 ]　ToolButtonのEnabledを設定する
//[ 引  数 ]　ボタン名と設定するEnabled
//*****************************************************************************
procedure TMainModule.SetActionEnabled(ActionName: string; Enabled: Boolean);
  function GetActionFromName(ActionName: string): TAction;
  var
    EditToolBar: TToolBar;
    i: Integer;
  begin
    Result := nil;

    EditToolBar := (BorlandIDEServices as INTAServices).ToolBar[TOOLBAR_NAME];
    if not Assigned(EditToolBar) then Exit;

    for i := 0 to High(FActions) - 1 do
    begin
      if FActions[i].Name = ActionName then
      begin
        Result := FActions[i];
        Exit;
      end;
    end;
  end;
var
  Action: TAction;
begin
  Action := GetActionFromName(ActionName);
  if Assigned(Action) then
    Action.Enabled := Enabled;
end;

//*****************************************************************************
//initialization/finalization
//*****************************************************************************
//*****************************************************************************
//登録
//*****************************************************************************
procedure Register;
begin
  MainModule := TMainModule.Create(nil);
//  OutputDebugLog('Register');
end;

//*****************************************************************************
//削除
//*****************************************************************************
procedure UnRegister;
type
  TToolbarArray = array of TWinControl;
var
  Toolbar: TToolBar;
  procedure RemoveFromToolbars(AToolBar: TToolBar);
  var
    ctx: TRttiContext;
    typ: TRttiType;
    fld: TRttiField;
    Toolbars: TToolbarArray;
    i, j: Integer;
    Button: TToolButton;
  begin
    typ := ctx.FindType('AppMain.TAppBuilder');
    if typ = nil then Exit;

    fld := typ.GetField('FToolbars');
    if fld = nil then Exit;

    fld.GetValue(Application.MainForm).ExtractRawData(@Toolbars);

    // ボタンの削除
    for i := Low(Toolbars) to High(Toolbars) do
    begin
      if Toolbars[i] = AToolBar then
        for j := TToolBar(Toolbars[i]).ButtonCount-1 downto 0 do
        begin
          Button := TToolBar(Toolbars[i]).Buttons[j];
          Toolbars[i].Perform(CM_CONTROLCHANGE, WPARAM(Button), LPARAM(False));
          Button.Free;
        end;
    end;

    // 追加したツールバーの除去
    for i := High(Toolbars) downto Low(Toolbars) do
      if Toolbars[i] = AToolBar then
        Delete(Toolbars, i, 1);

    TValue.From(Toolbars).ExtractRawData(PByte(Application.MainForm) + fld.Offset);
  end;
begin
  Toolbar := (BorlandIDEServices as INTAServices).ToolBar[TOOLBAR_NAME];
  RemoveFromToolbars(Toolbar);
  Toolbar.Free;

  MainModule.Free;
end;

//*****************************************************************************
//initialization/finalization
//*****************************************************************************
initialization
finalization
	UnRegister;
end.



