unit Module;

interface

uses
  Windows, Classes, Controls, ImgList, AppEvnts, ComCtrls, ToolsAPI, Dialogs, Forms,
  System.ImageList, System.Rtti, System.Actions, Vcl.ActnList;

procedure Register;

type
  TMainModule = class(TDataModule)
    IconList: TImageList;
    ApplicationEvents: TApplicationEvents;
    ActionList: TActionList;
    Action1: TAction;
    Action2: TAction;
    Action3: TAction;
    Action4: TAction;
    Action5: TAction;
    Action6: TAction;
    Action7: TAction;
    Action8: TAction;
    Action9: TAction;
    Action10: TAction;
    Action11: TAction;
    Action12: TAction;
    Action13: TAction;
    Action14: TAction;
    Action15: TAction;
    procedure ApplicationEventsIdle(Sender: TObject; var Done: Boolean);
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    FSearchText: string;
    procedure OnSetBookmarkClick(Sender: TObject);
    procedure OnNextBookmarkClick(Sender: TObject);
    procedure OnFindFormClick(Sender: TObject);
    procedure OnNextFindClick(Sender: TObject);
    procedure OnIndentClick(Sender: TObject);
    procedure OnCommentClick(Sender: TObject);
    procedure SetButtonEnabled(ButtonName: string; Enabled: Boolean);
  public
    procedure OnShowMethodFormClick(Sender: TObject);
    procedure Indent(Sign: Integer);
  end;

  TIDENotifier = class(TNotifierObject, IOTAIDENotifier)
  private
    procedure FileNotification(NotifyCode: TOTAFileNotification; const FileName: string; var Cancel: Boolean);
    procedure BeforeCompile(const Project: IOTAProject; var Cancel: Boolean); overload;
    procedure AfterCompile(Succeeded: Boolean); overload;
  end;
var
  MainModule: TMainModule;
  IDENo1: Integer;

implementation
{$R *.dfm}

uses
  FindForm, MethodForm, GlobalUnit;

const
  TOOLBAR_NAME = 'IDEEditToolBarEx';

//*****************************************************************************
//[ �T  �v ]�@Toolbar�Ƃ��̒���ToolButton�̍쐬���s��
//*****************************************************************************
procedure TMainModule.DataModuleCreate(Sender: TObject);
var
  IDEServices: INTAServices;
  ToolBar: TToolBar;
  ToolButton: TToolButton;
begin
  FSearchText := '';
  IDEServices := (BorlandIDEServices as INTAServices);
  ToolBar := IDEServices.NewToolbar(TOOLBAR_NAME, '���[�U��`�c�[���o�[');
  Assert(Assigned(ToolBar), '���[�U��`�̃c�[���o�[�̍쐬�Ɏ��s���܂���');
  ToolBar.Images := IconList;

  //��4������True�ɂ����Style=tbsDivider�ƂȂ邪�A��3�����̐ݒ肪�s�v�ɂȂ�
  //��4������False�ɂ����TSpeedButton���쐬�����݂�����
  ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME,'SetBookmark',nil ,True) as TToolButton;
  with ToolButton do
  begin
    Style      := tbsButton;
    OnClick    := OnSetBookmarkClick;
    ImageIndex := 0;
    Hint       := '������̐ݒ�/����';
  end;
  ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME,'NextBookmark',nil ,True) as TToolButton;
  with ToolButton do
  begin
    Style      := tbsButton;
    OnClick    := OnNextBookmarkClick;
    ImageIndex := 1;
    Hint       := '���̂�����';
  end;
  ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME,'PrevBookmark',nil ,True) as TToolButton;
  with ToolButton do
  begin
    Style      := tbsButton;
    OnClick    := OnNextBookmarkClick;
    ImageIndex := 2;
    Hint       := '���O�̂�����';
  end;

  //�Z�p���[�^
  IDEServices.AddToolButton(TOOLBAR_NAME,'Divider1',nil ,True);
  ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME,'FindForm',nil ,True) as TToolButton;
  with ToolButton do
  begin
    Style      := tbsButton;
    OnClick    := OnFindFormClick;
    ImageIndex := 3;
    Hint       := '����';
  end;
  ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME,'NextFind',nil ,True) as TToolButton;
  with ToolButton do
  begin
    Style      := tbsButton;
    OnClick    := OnNextFindClick;
    ImageIndex := 4;
    Hint       := '�������';
  end;
  ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME,'PrevFind',nil ,True) as TToolButton;
  with ToolButton do
  begin
    Style      := tbsButton;
    OnClick    := OnNextFindClick;
    ImageIndex := 5;
    Hint       := '�O������';
  end;

  //�Z�p���[�^
  IDEServices.AddToolButton(TOOLBAR_NAME,'Divider2',nil ,True);
  ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME,'Indent',nil ,True) as TToolButton;
  with ToolButton do
  begin
    Style      := tbsButton;
    OnClick    := OnIndentClick;
    ImageIndex := 6;
    Hint       := '�C���f���g';
  end;
  ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME,'Outdent',nil ,True) as TToolButton;
  with ToolButton do
  begin
    Style      := tbsButton;
    OnClick    := OnIndentClick;
    ImageIndex := 7;
    Hint       := '�C���f���g����';
  end;

  //�Z�p���[�^
  IDEServices.AddToolButton(TOOLBAR_NAME,'Divider3',nil ,True);
  ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME,'Comment',nil ,True) as TToolButton;
  with ToolButton do
  begin
    Style      := tbsButton;
    OnClick    := OnCommentClick;
    ImageIndex := 8;
    Hint       := '�R�����g��';
  end;
  ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME,'UnComment',nil ,True) as TToolButton;
  with ToolButton do
  begin
    Style      := tbsButton;
    OnClick    := OnCommentClick;
    ImageIndex := 9;
    Hint       := '�R�����g����';
  end;

  //�Z�p���[�^
  IDEServices.AddToolButton(TOOLBAR_NAME,'Divider4',nil ,True);
  ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME,'ShowMethods',nil ,True) as TToolButton;
  with ToolButton do
  begin
    Style      := tbsButton;
    OnClick    := OnShowMethodFormClick;
    ImageIndex := 10;
    Hint       := '���\�b�h�ꗗ';
  end;

//  ToolButton := IDEServices.AddToolButton(TOOLBAR_NAME,'ElseMenu',nil ,True) as TToolButton;
//  with ToolButton do
//  begin
//    Style := tbsButton;
//    DropdownMenu := ElseMenu;
//    ImageIndex := 11;
//  end;

  //�Ō�̃{�^���̔p�����Ȃ���΁A�{�^���̕\�����������Ȃ�Ȃ��̂Ń_�~�[���쐬���A���p��
  IDEServices.AddToolButton(TOOLBAR_NAME,'Dummy',nil ,True).Free;

  //�����ݒ�
  ToolBar.Visible := True;
  FindDlg := TFindDlg.Create(Self);
  MethodListForm := TMethodListForm.Create(Self);
end;

//*****************************************************************************
//[ �T  �v ]�@Toolbar��j������
//*****************************************************************************
procedure TMainModule.DataModuleDestroy(Sender: TObject);
var
  EditToolBar: TToolBar;
begin
  FindDlg.Free;
  MethodListForm.Free;
  EditToolBar := (BorlandIDEServices as INTAServices).ToolBar[TOOLBAR_NAME];
  if Assigned(EditToolBar) then EditToolBar.Free;
end;

//*****************************************************************************
//[�C�x���g]�@������̐ݒ�/�����N���b�N��
//[ �T  �v ]�@�J�[�\���s�̂�����̐ݒ�܂��͉������s��
//*****************************************************************************
procedure TMainModule.OnSetBookmarkClick(Sender: TObject);
var
  BookMarkID: Integer;
  Row: Integer;
begin
  if not CheckEditView() then Exit;

  //���݈ʒu
  Row := EditView.Position.Row;

  BookMarkID := GetActLineBookMarkID();
  if BookMarkID >= 0 then
  begin
    //������̉���
    EditView.BookmarkToggle(BookMarkID);
  end
  else
  begin
    //������̐ݒ�
    BookMarkID := GetUnusedBookMarkID();
    if BookMarkID < 10 then
    begin
      //�s�̐擪�ɂ������ݒ�
      EditView.Position.Move(Row, 1);
      EditView.BookmarkRecord(BookMarkID);
     end
    else
    begin
      MessageDlg('�������10�܂ł����ݒ�ł��܂���', mtInformation, [mbOK], 0);
    end;
  end;

  //�ĕ`��
  EditView.MoveViewToCursor;
  EditView.Paint;
end;

//*****************************************************************************
//[�C�x���g]�@���̂����� or ���O�̂�����N���b�N��
//[ �T  �v ]�@�J�[�\���s���O�܂��͒���̂�����ɃW�����v
//*****************************************************************************
procedure TMainModule.OnNextBookmarkClick(Sender: TObject);
var
  BookMarkID: Integer;
begin
  if not CheckEditView() then Exit;

  case (Sender as TComponent).Name[1] of
  'N': //NextBookmark
    begin
      //�����BookMarkID���擾
      BookMarkID := GetNextBookmarkID(+1);
    end;
  'P': //PrevBookmark
    begin
      //���O��BookMarkID���擾
      BookMarkID := GetNextBookmarkID(-1);
    end;
  else Exit;
  end;

  if BookMarkID >= 0 then
  begin
    EditView.BookmarkGoto(BookMarkID);
    EditView.MoveViewToCursor;

    //�ĕ`��
    EditView.Paint;
  end;
end;

//*****************************************************************************
//[�C�x���g]�@�����{�^���N���b�N��
//[ �T  �v ]�@�����t�H�[����\������
//*****************************************************************************
procedure TMainModule.OnFindFormClick(Sender: TObject);
var
  SearchResult: Boolean;
begin
  if not CheckEditView() then Exit;

  FindDlg.FindText := GetCursorWord();
  if FindDlg.ShowModal() = mrCancel then Exit;
//  FSearchText := WideStringToUtf8(FindDlg.FindText);
  FSearchText := FindDlg.FindText;

  with EditView.Position.SearchOptions do
  begin
    CaseSensitive     := FindDlg.chkCase.Checked;
    RegularExpression := FindDlg.chkRegEx.Checked;
    WordBoundary      := FindDlg.chkWord.Checked;
    SearchText        := FSearchText;
    WholeFile         := True;
    FromCursor        := True;
    Direction := sdForward;
  end;

  SearchResult := EditView.Position.SearchAgain();
  if SearchResult = False then
  begin
    EditView.Position.SearchOptions.Direction := sdBackward;
    SearchResult := EditView.Position.SearchAgain();
    if SearchResult = False then
    begin
      MessageDlg('������܂���ł���', mtInformation, [mbOK], 0);
      Exit;
    end;
  end;

  //�ĕ`��
  EditView.MoveViewToCursor;
  EditView.Paint;
end;

//*****************************************************************************
//[�C�x���g]�@������� or �O�������N���b�N��
//[ �T  �v ]�@�J�[�\���s���O�܂��͒���̌��������ɃW�����v
//*****************************************************************************
procedure TMainModule.OnNextFindClick(Sender: TObject);
var
  SearchResult: Boolean;
  TopRow, LeftCol: Integer;
begin
  if not CheckEditView() then Exit;

//  �o�O���ۂ��̂Ŏg���Ȃ�����
//  SearchText := EditView.Position.SearchOptions.SearchText;

  if FSearchText = '' then Exit;

  with EditView.Position.SearchOptions do
  begin
    CaseSensitive     := FindDlg.chkCase.Checked;
    RegularExpression := FindDlg.chkRegEx.Checked;
    WordBoundary      := FindDlg.chkWord.Checked;
    SearchText        := FSearchText;
    WholeFile         := True;
    FromCursor        := True;
  end;

  //��ʏ�Ԃ̕ۑ�
  TopRow  := EditView.TopRow;
  LeftCol := EditView.LeftColumn;
  EditView.Block.Save;
  EditView.Position.Save;

  case (Sender as TComponent).Name[1] of
  'N': //NextFind
    begin
      EditView.Position.MoveRelative(0, 1);
      EditView.Position.SearchOptions.Direction := sdForward;
    end;
  'P': //PrevFind
    begin
      EditView.Position.MoveRelative(0,-1);
      EditView.Position.SearchOptions.Direction := sdBackward;
    end;
  end;

  SearchResult := EditView.Position.SearchAgain();
  if SearchResult = False then
  begin
    //�܂�Ԃ��Č���
    if EditView.Position.SearchOptions.Direction = sdForward then
        EditView.Position.Move(1, 1)
    else
        EditView.Position.MoveEOF;

    SearchResult := EditView.Position.SearchAgain();
    if SearchResult = False then
    begin
      //��ʏ�Ԃ̕���
      EditView.Position.Restore;
      EditView.Block.Restore;
      EditView.SetTopLeft(TopRow, LeftCol);
      MessageDlg('������܂���ł���', mtInformation, [mbOK], 0);
      Exit;
    end
    else
    begin
      //�ĕ`��
      EditView.MoveViewToCursor;
      EditView.Paint;
      MessageDlg('�܂�Ԃ��܂���', mtInformation, [mbOK], 0);
      Exit;
    end;
  end;

  //�ĕ`��
  EditView.MoveViewToCursor;
  EditView.Paint;
end;

//*****************************************************************************
//[�C�x���g]�@�C���f���g or �C���f���g�����N���b�N��
//[ �T  �v ]�@�I���s���������܂��͎��グ
//*****************************************************************************
procedure TMainModule.OnIndentClick(Sender: TObject);
begin
  if not CheckEditView() then Exit;

  case (Sender as TComponent).Name[1] of
  'I':Indent(+1);
  'O':Indent(-1);
  else Exit;
  end;

  //�ĕ`��
  EditView.Paint;
end;

//*****************************************************************************
//[�C�x���g]�@�C���f���g or �C���f���g����
//[ ��  �� ]�@+1:�C���f���g�A-1:�C���f���g����
//[ �T  �v ]�@�I���s���������܂��͎��グ
//*****************************************************************************
procedure TMainModule.Indent(Sign: Integer);
var
  Indent: Integer;
  StartingRow, EndingRow: Integer;
  MinusRow: Integer;
begin
  //�u���b�N�I���ŃJ�[�\�����s���Ȃ炻�̍s�͑ΏۊO�ɂ���
  if (EditView.Block.Style <> btLine) and
     (EditView.Block.Size <> 0) and
     (EditView.Block.EndingColumn = 1) then MinusRow := 1
                                       else MinusRow := 0;
  //�I�����s�I���ɐ؂�ւ�
  EditView.Block.Style := btLine;
  StartingRow  := EditView.Block.StartingRow;
  EndingRow    := EditView.Block.EndingRow - MinusRow;

  SelectRows(StartingRow, EndingRow);

  //�C���f���g�����s
  Indent := EditView.Buffer.EditOptions.BlockIndent;
  EditView.Block.Indent(Indent * Sign);
end;

//*****************************************************************************
//[�C�x���g]�@�R�����g�� or �R�����g�����N���b�N��
//[ �T  �v ]�@�I���s���R�����g���܂��͉���
//*****************************************************************************
procedure TMainModule.OnCommentClick(Sender: TObject);
var
  i: Integer;
  StartingRow, EndingRow: Integer;
  MinusRow: Integer;
begin
  if not CheckEditView() then Exit;

  //�u���b�N�I���ŃJ�[�\�����s���Ȃ炻�̍s�͑ΏۊO�ɂ���
  if (EditView.Block.Style <> btLine) and
     (EditView.Block.Size <> 0) and
     (EditView.Block.EndingColumn = 1) then MinusRow := 1
                                       else MinusRow := 0;
  //�I�����s�I���ɐ؂�ւ�
  EditView.Block.Style := btLine;
  StartingRow  := EditView.Block.StartingRow;
  EndingRow    := EditView.Block.EndingRow - MinusRow;

  for i := StartingRow to EndingRow do
  begin
    EditView.Position.Move(i, 1);
    case (Sender as TComponent).Name[1] of
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

  //�R�����g�s��I��
  SelectRows(StartingRow, EndingRow);

  //�ĕ`��
  EditView.Paint;
end;

//*****************************************************************************
//[�C�x���g]�@���\�b�h�ꗗ�N���b�N��
//[ �T  �v ]�@���\�b�h�ꗗ�t�H�[����\������
//*****************************************************************************
procedure TMainModule.OnShowMethodFormClick(Sender: TObject);
begin
  if not CheckEditView() then Exit;

  MethodListForm.ShowModal;
end;

//*****************************************************************************
//[�C�x���g]�@�A�v���P�[�V�������A�C�h���̎�
//[ �T  �v ]�@�eToolButton��Enabled��ݒ肷��
//*****************************************************************************
procedure TMainModule.ApplicationEventsIdle(Sender: TObject; var Done: Boolean);
var
  UsedBookMarkCount: Integer;
//  ST: SYSTEMTIME;
begin
//  GetLocalTime(ST);
//  OutputDebugString(PChar(Format('���� %2d:%2d:%2d.%3d    Name:%s   ClassName:%s',
//   [ST.wHour, ST.wMinute, ST.wSecond, ST.wMilliseconds,
//    TComponent(Sender).Name, Sender.ClassType.ClassName])));

  if not CheckEditView() then
  begin
    SetButtonEnabled('SetBookmark' , False);
    SetButtonEnabled('NextBookmark', False);
    SetButtonEnabled('PrevBookmark', False);
    SetButtonEnabled('FindForm'    , False);
    SetButtonEnabled('NextFind'    , False);
    SetButtonEnabled('PrevFind'    , False);
    SetButtonEnabled('Indent'      , False);
    SetButtonEnabled('Outdent'     , False);
    SetButtonEnabled('Comment'     , False);
    SetButtonEnabled('UnComment'   , False);
    SetButtonEnabled('ShowMethods' , False);
    Exit;
  end;

  SetButtonEnabled('NextFind', (FSearchText <> ''));
  SetButtonEnabled('PrevFind', (FSearchText <> ''));

  UsedBookMarkCount := GetUsedBookMarkCount();
  SetButtonEnabled('NextBookmark', (UsedBookMarkCount > 0));
  SetButtonEnabled('PrevBookmark', (UsedBookMarkCount > 0));

  SetButtonEnabled('SetBookmark', True);
  SetButtonEnabled('FindForm'   , True);
  SetButtonEnabled('Indent'     , True);
  SetButtonEnabled('Outdent'    , True);
  SetButtonEnabled('Comment'    , True);
  SetButtonEnabled('UnComment'  , True);
  SetButtonEnabled('ShowMethods', True);
end;

//*****************************************************************************
//[ �T  �v ]�@ToolButton��Enabled��ݒ肷��
//[ ��  �� ]�@�{�^�����Ɛݒ肷��Enabled
//*****************************************************************************
procedure TMainModule.SetButtonEnabled(ButtonName: string; Enabled: Boolean);
  function GetButtonFromName(ButtonName: string): TToolButton;
  var
    EditToolBar: TToolBar;
    i: Integer;
    ToolButton: TToolButton;
  begin
    Result := nil;

    EditToolBar := (BorlandIDEServices as INTAServices).ToolBar[TOOLBAR_NAME];
    if not Assigned(EditToolBar) then Exit;

    for i := 0 to EditToolBar.ButtonCount - 1 do
    begin
      ToolButton := EditToolBar.Buttons[i];
      if ToolButton.Name = ButtonName then
      begin
        Result := ToolButton;
        Exit;
      end;
    end;
  end;
var
  ToolButton: TToolButton;
begin
  ToolButton := GetButtonFromName(ButtonName);
  if Assigned(ToolButton) then
    ToolButton.Enabled := Enabled;
end;

//*****************************************************************************
//initialization/finalization
//*****************************************************************************
{ TIDENotifier }
//*****************************************************************************
//[�C�x���g] �t�@�C���I�[�v����N���[�Y��
//[ �T  �v ]
//*****************************************************************************
procedure TIDENotifier.FileNotification(NotifyCode: TOTAFileNotification;
  const FileName: string; var Cancel: Boolean);
var
  ST: SYSTEMTIME;
  Notify: string;
begin
//  case NotifyCode of
//    ofnFileOpening:          Notify := 'ofnFileOpening';
//    ofnFileOpened:           Notify := 'ofnFileOpened';
//    ofnFileClosing:          Notify := 'ofnFileClosing';
//    ofnDefaultDesktopLoad:   Notify := 'ofnDefaultDesktopLoad';
//    ofnDefaultDesktopSave:   Notify := 'ofnDefaultDesktopSave';
//    ofnProjectDesktopLoad:   Notify := 'ofnProjectDesktopLoad';
//    ofnProjectDesktopSave:   Notify := 'ofnProjectDesktopSave';
//    ofnPackageInstalled:     Notify := 'ofnPackageInstalled';
//    ofnPackageUninstalled:   Notify := 'ofnPackageUninstalled';
//    ofnActiveProjectChanged: Notify := 'ofnActiveProjectChanged';
//    ofnProjectOpenedFromTemplate:
//                             Notify := 'ofnProjectOpenedFromTemplate';
//    else                     Notify := 'other';
//  end;
//  GetLocalTime(ST);
//  Clipboard.AsText := Format('%2d:%2d:%2d.%3d %s %s',
//   [ST.wHour, ST.wMinute, ST.wSecond, ST.wMilliseconds, Notify,FileName]);
end;

//*****************************************************************************
//[�C�x���g] Before�R���p�C����
//[ �T  �v ] �{���͂����Ŏ������������A�z��O�̃^�C�~���O�ŃC�x���g����
//*****************************************************************************
procedure TIDENotifier.BeforeCompile(const Project: IOTAProject; var Cancel: Boolean);
var
  ST: SYSTEMTIME;
begin
//  GetLocalTime(ST);
//  Clipboard.AsText := Format('BeforeCompile %2d:%2d:%2d.%3d',
//   [ST.wHour, ST.wMinute, ST.wSecond, ST.wMilliseconds]);
end;

//*****************************************************************************
//[�C�x���g] After�R���p�C����
//[ �T  �v ] �{���͂����Ŏ������������A�z��O�̃^�C�~���O�ŃC�x���g����
//*****************************************************************************
procedure TIDENotifier.AfterCompile(Succeeded: Boolean);
var
  ST: SYSTEMTIME;
begin
//  GetLocalTime(ST);
//  Clipboard.AsText := Format('AfterCompile %2d:%2d:%2d.%3d',
//   [ST.wHour, ST.wMinute, ST.wSecond, ST.wMilliseconds]);
end;

//*****************************************************************************
//�o�^
//*****************************************************************************
procedure Register;
begin
  IDENo1 := (BorlandIDEServices as IOTAServices).AddNotifier(TIDENotifier.Create);

  MainModule := TMainModule.Create(nil);
end;

//*****************************************************************************
//�폜
//*****************************************************************************
procedure UnRegister;
type
  TToolbarArray = array of TWinControl;
var
  toolbar: TToolBar;
  i: Integer;
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

    // �{�^���̍폜
    for i := Low(Toolbars) to High(Toolbars) do
    begin
      if Toolbars[i] = AToolBar then
        for j := TToolBar(Toolbars[i]).ButtonCount-1 downto 0 do
        begin
          Button := TToolBar(Toolbars[i]).Buttons[j];
          OutputDebugString(Pchar(Button.Name));
          Toolbars[i].Perform(CM_CONTROLCHANGE, WPARAM(Button), LPARAM(False));
          Button.Free;
        end;
    end;

    // �ǉ������c�[���o�[�̏���
    for i := High(Toolbars) downto Low(Toolbars) do
      if Toolbars[i] = AToolBar then
        Delete(Toolbars, i, 1);

    TValue.From(Toolbars).ExtractRawData(PByte(Application.MainForm) + fld.Offset);
  end;
begin
  toolbar := (BorlandIDEServices as INTAServices).ToolBar[TOOLBAR_NAME];
  RemoveFromToolbars(toolbar);
  toolbar.Free;

  (BorlandIDEServices as IOTAServices).RemoveNotifier(IDENo1);
  MainModule.Free;
end;

//*****************************************************************************
//initialization/finalization
//*****************************************************************************
initialization
finalization
	UnRegister;
end.

