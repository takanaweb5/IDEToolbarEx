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
    FimplementationRow: Integer;//implementation�̂���s
    FinterfaceRow: Integer; //interface�̊J�n�s
    FInfoIndex: Integer;    //FMethodList�̋N�����̃J�[�\���ʒu�̃C���f�b�N�X
    FTopRow: Integer;       //�t�H�[���N�����̉�ʂ̐擪�s
    FEditPos: TOTAEditPos;  //�t�H�[���N�����̃J�[�\���ʒu
    FText: string;          //�ҏW���̃t�@�C���̂��ׂĂ̍s
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
//[ �T  �v ]�@�T�C�Y�O���b�v���쐬����
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

  //�T�u�N���X��
  FOrgWndProc := Parent.WindowProc;
  Parent.WindowProc := SubClassProc;
end;

//*****************************************************************************
//[ �T  �v ]�@�f�X�g���N�^
//*****************************************************************************
destructor TSizeGrip.Destroy;
begin
  if (Parent <> nil) and Parent.HandleAllocated then
    Parent.WindowProc := FOrgWndProc;
  inherited Destroy;
end;

//*****************************************************************************
//[ �T  �v ]�@�t�H�[���̃T�u�N���X��
// �t�H�[���ړ����A�T�C�Y�O���b�v���m���ɍĕ`�悳����@
// �t�H�[���̍ő剻���A�T�C�Y�O���b�v��\�������Ȃ��@
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
//[�C�x���g]�@�T�C�Y�O���b�v�̕`�揈��
//*****************************************************************************
procedure TSizeGrip.Paint();
begin
  with Self.Canvas do
  begin
    DrawFrameControl(Handle, ClipRect, DFC_SCROLL, DFCS_SCROLLSIZEGRIP);
  end;
end;

//*****************************************************************************
//[�C�x���g]�@�T�C�Y�O���b�v��MouseDown��
//[ �T  �v ]�@�T�C�Y�̕ύX
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
//[�R���X�g���N�^]�@
//[ �T  �v ]�@TListBox��TListBoxEx�ɕϐg������
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
      //�v���p�e�B�̐��������[�v
      for i := 0 to Count - 1 do
      begin
        PropInfo := PropList[i]^;
        //�v���p�e�B���C�x���g�̎�
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
    Original.Name := ''; //���ꂪ�Ȃ��Əd����O������
    MemStream.Position := 0;
    MemStream.ReadComponent(Self);  //�C�x���g�̓R�s�[����Ȃ�

    //�C�x���g�̃R�s�[
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
//[ �T  �v ] ���X�g�{�b�N�X�̑I���s�̃t�H�[�J�X�̎l�p��`�悳���Ȃ��@
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
//[�C�x���g]�@���X�g�{�b�N�X�}�E�X�z�C�[����
//[ �T  �v ]�@���X�g�{�b�N�X���C�����悭�X�N���[��������
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
//[ �T  �v ]�@�\�[�g�p�̃��[�U��`�֐�
//*****************************************************************************
function StringListCompareStrings(List: TStringList; Index1, Index2: Integer): Integer;
var
  Info1, Info2: TLineInfo;
begin
  Info1 := MethodListForm.FMethodList[Integer(List.Objects[Index1])];
  Info2 := MethodListForm.FMethodList[Integer(List.Objects[Index2])];

  //�N���X���Ń\�[�g
  Result := AnsiCompareText(Info1.ClassName, Info2.ClassName);
  if Result <> 0 then Exit;

  //���\�b�h���Ń\�[�g
  Result := AnsiCompareText(Info1.MethodName, Info2.MethodName);
  if Result <> 0 then Exit;

  //�s�ԍ��Ń\�[�g
  Result := Info1.LineNo - Info2.LineNo;
end;

{ TMethodListForm }
//*****************************************************************************
//[�C�x���g]�@�A�v���P�[�V�������A�C�h���̎�
//[ �T  �v ]�@�e�R���g���[����Enabled��ݒ肷��
//*****************************************************************************
procedure TMethodListForm.ApplicationEventsIdle(Sender: TObject; var Done: Boolean);
begin
try
  //���t�H�[���ȊO��ApplicationEventsIdle�����s�����Ȃ�
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
//[�C�x���g]�@�t�H�[���쐬��
//[ �T  �v ]�@��������
//*****************************************************************************
procedure TMethodListForm.FormCreate(Sender: TObject);
begin
  //�T�C�Y�O���b�v�̍쐬
  TSizeGrip.Create(Self);

  //���X�g�{�b�N�X��ϐg������
  ListBox := TListBoxEx.CreateClone(ListBox);
  ListBox.ItemHeight := ListBox.ItemHeight + 4;

  //�t�H�[���̍ŏ����̐ݒ�
  Constraints.MinWidth := 420;
end;

//*****************************************************************************
//[�C�x���g]�@�t�H�[���\����
//[ �T  �v ]�@��ʂ̏����ݒ�
//*****************************************************************************
procedure TMethodListForm.FormShow(Sender: TObject);
  function GetLineIndex(LineNo: Integer): Integer;
  var
    i: Integer;
  begin
    Result := -1;

    //LineNo���܂ރ��\�b�h������
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

  //ApplicationEventsIdle��L���ɂ���
  ApplicationEvents.OnIdle := ApplicationEventsIdle;
  ApplicationEvents.Activate;

  //�N�����̃J�[�\���ʒu����ۑ�
  FTopRow  := EditView.TopRow;
  FEditPos := EditView.CursorPos;

  //�ҏW���̃t�@�C�������t�H�[���̃^�C�g���o�[�ɕ\��
  Caption := '���\�b�h�ꗗ - ' + ExtractFileName(EditView.Buffer.FileName);

  FText := GetAllText();

  //�e�R���g���[����Enabled�̐ݒ�
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

  //�J�[�\���ʒu�̃��\�b�h��I��
  LineNo := EditView.Position.Row;
  if GetLineIndex(LineNo) <> -1 then
  begin
    ListBox.ItemIndex := GetLineIndex(LineNo);
    FInfoIndex := Integer(ListBox.Items.Objects[ListBox.ItemIndex]);
    FMethodList[FInfoIndex].Select := True;
  end
  else
  begin
    //�擪��I��
    ListBox.ItemIndex := 0;
    FInfoIndex := -1;
  end;
end;

//*****************************************************************************
//[�C�x���g]�@�t�H�[����Close���鎞
//*****************************************************************************
procedure TMethodListForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  case ModalResult of
  mrOk : ;//�������Ȃ�
//   begin
//    //�I�����ꂽ���\�b�h�̃��W�b�N�J�n�ʒu�փJ�[�\�����ړ�
//    LineNo := GetLineInfo(ListBox.ItemIndex).LineNo;
//    LineNo := GetMethodLineInfo(LineNo, FText).LogicStart;
//    EditView.Position.Move(LineNo, 1);
//   end;
  mrCancel :
   begin
    //�t�H�[���N�����̃J�[�\���ʒu��\��
    EditView.SetTopLeft(FTopRow, 1);
    EditView.SetCursorPos(FEditPos);
   end;
  end;

  EditView.Paint;

  //�E�B���h�E�̃T�C�Y�����ɖ߂�
//  ShowWindow(Handle, SW_HIDE);
  Self.WindowState := wsNormal;

  FMethodList := nil;

  //ApplicationEventsIdle�𖳌��ɂ���
  ApplicationEvents.OnIdle := nil;
end;

//*****************************************************************************
//[�C�x���g]�@�t�H�[����Close����O�Ɋm�F����
//*****************************************************************************
procedure TMethodListForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if ModalResult = mrOk then Exit;

  if (FTopRow <> EditView.TopRow) or (FEditPos.Line <> EditView.CursorPos.Line) then
  begin
    case MessageDlg('���̈ʒu�ɖ߂��܂����H', mtInformation,
                                      [mbYes, mbNo, mbCancel], 0) of
    mrYes   : ModalResult := mrCancel;
    mrNo    : ModalResult := mrOk;
    mrCancel: CanClose := False;
    end;
  end;
end;

//*****************************************************************************
//[�C�x���g]�@KeyDown��
//[ �T  �v ] ESC�Ńt�H�[�������@
//*****************************************************************************
procedure TMethodListForm.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_ESCAPE then ModalResult := mrCancel;
end;

//*****************************************************************************
//[�C�x���g]�@����{�^���N���b�N��
//*****************************************************************************
procedure TMethodListForm.cmdOKClick(Sender: TObject);
begin
  ModalResult := mrOk;
end;

//*****************************************************************************
//[�C�x���g]�@���X�g�{�b�N�XKeyPress��
//[ �T  �v ]�@�@�X�y�[�X�L�[�Ńu�b�N�}�[�N�̐ݒ�/�������s��
//            �@��KeyDown�ł̓f�t�H���g�̓����}���ł��Ȃ��̂�KeyPress���g�p
//            �A�A���t�@�x�b�g�̃L�[�ŁA�Y�����郁�\�b�h������
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
        //���\�b�h��1�����ڂ��ƍ�(MethodName=''�̏ꍇ����O�Ȃ�)
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
  //���X�g�{�b�N�X���I������Ă��Ȃ��ƑΏۊO
  Index := ListBox.ItemIndex;
  if Index < 0 then Exit;

  //�X�y�[�X�L�[�Ńu�b�N�}�[�N�̐ݒ�/�������s��
  if Key = ' ' then
  begin
    ListBoxClickBookMark(Index);
    //�f�t�H���g�̑����}������
    Key := Chr(0);
    Exit;
  end;

  //�A���t�@�x�b�g�̃L�[�ŁA�Y�����郁�\�b�h������
  case Key of
   'A'..'Z','a'..'z','0'..'9','_':
    begin
      k := Integer(ListBox.Items.Objects[Index]);

      blnSerch := (FMethodList[k].Command <> C_CLASSDEFINE);
      if blnSerch = False then
        if (cmbCommand.Text = '���ׂ�') and (chkSort.Checked = True) then
          blnSerch := True;
        if chkClassName.Checked then
          blnSerch := True;

      if blnSerch then
      begin
        //�Y������N���X��Key�Ŏn�܂郁�\�b�h��I�����ꂽIndex�ȍ~�Ō���
        j := SerchMethod(FMethodList[k].ClassName, Index + 1);
        if j < 0 then
          //�I�����ꂽIndex�ȍ~�Ɍ�����Ȃ���΍ŏ����猟��
          j := SerchMethod(FMethodList[k].ClassName, 0);

        if j >= 0 then
        begin
          ListBox.ItemIndex := j;
          JumpSelectMethod(j);
        end;
      end;

      //�f�t�H���g�̑����}������
      Key := Chr(0);
    end;
  end;
end;

//*****************************************************************************
//[�C�x���g] ���X�g�{�b�N�XKeyDown��
//[ �T  �v ] DELETE�Ńu�b�N�}�[�N���폜�@
//*****************************************************************************
//procedure TMethodListForm.ListBoxKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
//begin
//  if Key = VK_DELETE then
//  begin
//  end;
//end;

//*****************************************************************************
//[ �T  �v ] ���X�g�{�b�N�X�őI�����ꂽ���\�b�h���̃u�b�N�}�[�N�����ׂč폜����
//[ ��  �� ] ���X�g�{�b�N�X�őI������Ă���Index
//*****************************************************************************
procedure TMethodListForm.ClearBookmark(Index: Integer);
var
  i, k: Integer;
  BMLineNo: Integer;
begin
  k := Integer(ListBox.Items.Objects[Index]);
  if FMethodList[k].BookMark = [] then Exit;
  FMethodList[k].BookMark := [];

  //�u�b�N�}�[�N�����[�v
  for i := 0 to 9 do
  begin
    BMLineNo := EditView.BookmarkPos[i].Line;
    if BMLineNo = 0 then Continue;
    if GetMethodLineInfo(BMLineNo, FText).StartRow = FMethodList[k].LineNo then
    begin
      //Bookmark�̉���
      EditView.BookmarkGoto(i);
      EditView.BookmarkToggle(i);
    end;
  end;

  //�ĕ`��
  EditView.Paint;
  ListBox.Repaint;
end;


//*****************************************************************************
//[ �T  �v ] ���X�g�{�b�N�X�̍��ڂ̕`���Ǝ��ɍs���@
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
    //�u�b�N�}�[�N�A�C�R���G���A�̕`��
    //*************************************
    if Index = Items.Count - 1 then
      //�Ō�̍s�̎��́A�c��̗̈���Ώ�
      R := Rect(0, ARect.Top, LEFTWIDTH, ClientRect.Bottom)
    else
      R := Rect(0, ARect.Top, LEFTWIDTH, ARect.Bottom);
    Canvas.Brush.Color := RGB(244, 244, 244);
    Canvas.FillRect(R);
    Canvas.Pen.Color := RGB(153, 153, 204);
    Canvas.MoveTo(LEFTWIDTH, ARect.Top);
    Canvas.LineTo(LEFTWIDTH, R.Bottom);

    //�J�����g�s�́A!�}�[�N������
    if GetLineInfo(Index).Select then
    begin
      Canvas.Font.Color := clRed;
      Canvas.TextOut(0, R.Top + 1, '!');
    end;

    //�u�b�N�}�[�N�̃A�C�R��������
    if bsOn in GetLineInfo(Index).BookMark then
      ImageList.Draw(Canvas,3, ARect.Top, 0)
    else if bsGray in GetLineInfo(Index).BookMark then
      ImageList.Draw(Canvas,3, ARect.Top, 1);

    //*************************************
    //��L�ȊO�̃G���A�̕`��
    //*************************************
    R := Rect(ARect.Left + LEFTWIDTH + 1, ARect.Top, ARect.Right, ARect.Bottom);

    Canvas.Brush.Color := GetLineInfo(Index).Color;

    if cmbCommand.Text = '�N���X�錾' then
        Canvas.Brush.Color := $FFDDFE
    else
      if (chkSort.Checked = False) and (chkClassName.Checked = False) and
         (GetLineInfo(Index).Command = C_CLASSDEFINE) then
        Canvas.Brush.Color := $FFDDFE;

    //�t�H�[�J�X�̂���s�̐F��ύX����
    if odSelected in State then
      Canvas.Brush.Color := Canvas.Brush.Color - $0C0C0C;

    Canvas.FillRect(R);
    Canvas.Font.Color := clBlack;
    Canvas.TextOut(R.Left + 6, R.Top + 1, Items[Index]);

    //�t�H�[�J�X�̓_���g
    if odSelected in State then Canvas.DrawFocusRect(R);

    //implementation���̊J�n�O�ɉ���������
    if (chkSort.Checked = False) and (chkClassName.Checked = False) then
    begin
      if (Index > 0) and
         (GetLineInfo(Index - 1).LineNo < FimplementationRow) and
         (FimplementationRow < GetLineInfo(Index).LineNo) then
      begin
        //�ɍׂ̓_����`��
        DrawDotLine(Canvas, LEFTWIDTH + 1, ARect.Right);
//        Canvas.Pen.Color := clBlack;
//        Canvas.MoveTo(LEFTWIDTH + 1, ARect.Top);
//        Canvas.LineTo(ARect.Right, ARect.Top);
      end;
    end;
  end;
end;

//*****************************************************************************
//[�C�x���g]�@���X�g�{�b�N�X�N���b�N��
//[ �T  �v ]�@�I�����ꂽ���\�b�h��\������
//*****************************************************************************
procedure TMethodListForm.ListBoxClick(Sender: TObject);
begin
  JumpSelectMethod(ListBox.ItemIndex);
end;

//*****************************************************************************
//[ �T  �v ]�@���X�g�{�b�N�̑I�����ꂽ���\�b�h��\������
//[ ��  �� ]�@ListBoxIndex
//*****************************************************************************
procedure TMethodListForm.JumpSelectMethod(Index: Integer);
var
  LineNo: Integer;
  MethodLineInfo: TMethodLineInfo;
begin
  LineNo := GetLineInfo(Index).LineNo;

  if GetLineInfo(Index).Command = C_CLASSDEFINE then
  begin
    //�ĕ`��
    EditView.SetTopLeft(GetCommentStart(LineNo, FText), 1);
    EditView.Paint;
    EditView.Position.Move(LineNo, 1);
  end
  else
  begin
    MethodLineInfo := GetMethodLineInfo(LineNo, FText);
    if MethodLineInfo.StartRow = 0 then Exit;

    //�ĕ`��
    EditView.SetTopLeft(MethodLineInfo.CommentStart, 1);
    EditView.Paint;
    EditView.Position.Move(MethodLineInfo.LogicStart, 1);
  end;
end;

//*****************************************************************************
//[�C�x���g]�@���X�g�{�b�N�X�_�u���N���b�N��
//*****************************************************************************
procedure TMethodListForm.ListBoxDblClick(Sender: TObject);
begin
  ModalResult := mrOk;
end;

//*****************************************************************************
//[�C�x���g]�@���X�g�{�b�N�X�}�E�X�_�E����
//[ �T  �v ]�@�N���b�N���ꂽ�̈�Ő����ς���
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
//[�C�x���g]�@�u�b�N�}�[�N�A�C�R���G���A�̃N���b�N��
//[ �T  �v ]�@�u�b�N�}�[�N�A�C�R���̐ݒ�
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
    //���łɐݒ肳��Ă���Ή���
    Exclude(FMethodList[k].BookMark, bsOn)
  else
  begin
    //���ݒ�ł���ΐݒ�
    if GetUnusedBookMarkID() >= 10 then
    begin
      MessageDlg('�������10�܂ł����ݒ�ł��܂���', mtInformation, [mbOK], 0);
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
    //������̐ݒ�
    BookMarkID := GetUnusedBookMarkID();
    EditView.Position.Move(MethodLineInfo.StartRow, 1);
    EditView.BookmarkRecord(BookMarkID);
  end
  else
  begin
    //�u�b�N�}�[�N�����[�v
    for i := 0 to 9 do
    begin
      BMLineNo := EditView.BookmarkPos[i].Line;
      if BMLineNo = 0 then Continue;
      if (BMLineNo = MethodLineInfo.StartRow) or
         (BMLineNo = MethodLineInfo.LogicStart) then
      begin
        //Bookmark�̉���
        EditView.BookmarkGoto(i);
        EditView.BookmarkToggle(i);
      end;
    end;
  end;

  //�ĕ`��
  EditView.Paint;
  ListBox.Repaint;

  //�_�u���N���b�N�C�x���g��}�~
  ListBox.OnDblClick := nil;
end;

//*****************************************************************************
//[�C�x���g]�@�u�b�N�}�[�N�A�C�R���G���A�ȊO�̃N���b�N��
//[ �T  �v ]�@�_�u���N���b�N�C�x���g�𕜊�
//*****************************************************************************
procedure TMethodListForm.ListBoxClickList(Index: Integer);
begin
  ListBox.OnDblClick := ListBoxDblClick;
end;

//*****************************************************************************
//[�C�x���g]  �t�H�[���\����
//          �@�\�[�g�N���b�N��
//          �@�`�F�b�N�{�b�N�X�`�F�b�N��
//[ �T  �v ]�@FMethodList�̓��e���烊�X�g�{�b�N�X��ݒ肷��
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

  //�I������Ă���s�̏����擾����
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

  //�Ώۂ̍s��I������
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
        if (cmbCommand.Text <> '���ׂ�') and (cmbCommand.Text <> '�N���X�錾')  then
          Continue;

        if cmbCommand.Text = '�N���X�錾' then
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

    //�\�[�g���`�F�b�N����Ă���΃\�[�g������
    if chkSort.Checked then ST.CustomSort(StringListCompareStrings);

    ListBox.Items.Assign(ST);
  finally
    ST.Free;
  end;

  //�I������Ă����s���đI��������
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

  //�w�i�F�̏���ݒ肷��
  ClearColor();
  SetColor(True);

  //�I���s�̔w�i�F���I��O�ƕς������I��O�Ɠ����ɂȂ�悤�ɂ�蒼��
  if ListBox.ItemIndex >= 0 then
  begin
    if (iColor <> clWhite) and
       (iColor <> GetLineInfo(ListBox.ItemIndex).Color) then
    begin
      //�w�i�F�̏����Đݒ肷��
      SetColor(False);
    end;
  end;

  //�ĕ`��
  ListBox.Repaint;
  ListBox.SetFocus;
end;

//*****************************************************************************
//[ �T  �v ]�@�w�i�F��ݒ肷��
//[ ��  �� ]�@�ŏ��̐F��������ɂ��邩�̃g�O��
// �N���X���������l�̎��́u���v
// �N���X���������Ԃ́A����F�œ���
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

    //���O�̃N���X���ƈႤ���A�F��ύX
    if AnsiSameText(GetLastName(i), GetLineInfo(i).ClassName) = False then
      ColorToggle := not ColorToggle;

    FMethodList[j].Color := DrawingColor[ColorToggle];
  end;
end;

//*****************************************************************************
//[ �T  �v ]�@FMethodList��ݒ肷��
//[ ��  �� ]�@�ҏW�t�@�C���̑S�s
//[ �߂�l ]�@(procedure or function) �̍s��
//*****************************************************************************
function TMethodListForm.SetMethodList(const Text: string): Integer;
begin
  //LineNo/Command/Define/ClassName/MethodName ��ݒ�
  Result := SetMethodLines(Text);
  if Result = 0 then Exit;

  //FMethodList��Color���N���A����
  ClearColor();

  //FMethodList��BookMark��ݒ肷��
  SetBookMarkInfo();
end;

//*****************************************************************************
//[ �T  �v ]�@FMethodList�� Color(�w�i�F)���N���A����
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
//[ �T  �v ]�@FMethodList�� LineNo/Command/Define/ClassName/MethodName ��ݒ�
//[ ��  �� ]�@�ҏW�t�@�C���̑S�s
//[ �߂�l ]�@(procedure or function) �̍s��
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

    //�s����Index�����킹�邽�ߐ擪��1�s�}��
    ST.Insert(0, '');

    //�s�������[�v
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

    //�Y���s���̎擾
    j := 0;
    //�s�������[�v
    RegExp.Pattern := C_CLASS;
    for i := FinterfaceRow + 1 to ST.Count - 1 do
    begin
      if RegExp.Test(ST[i]) then
      begin
        TmpArray[i] := 1;
        Inc(j);
      end;
    end;

    //�s�������[�v(implementation��)
    RegExp.Pattern := C_METHOD;
    for i := FimplementationRow + 1 to ST.Count - 1 do
    begin
      //forward�錾�͑ΏۊO
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
    //�s�������[�v���AFMethodList[j]��ݒ�
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
          ClassName := Match.SubMatches[0]; //��:TMethodListForm
          Define  := Match.SubMatches[1]; //��:class(TForm)
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
          Command := Match.SubMatches[0]; //��:procedure
          Define  := Match.SubMatches[1]; //��:TMethodListForm.SetMethodList(const Text: string);

          if Match.SubMatches[5] = '' then
          begin
            //��:TMethodListForm.SetMethodList(const Text: string);
            ClassName  := Match.SubMatches[3]; //��:TMethodListForm
            MethodName := Match.SubMatches[4]; //��:SetMethodList
          end
          else
          begin
            //��:StringListCompareStrings(Index1, Index2: Integer): Integer;
            ClassName  := '';
            MethodName := Match.SubMatches[5]; //��:StringListCompareStrings
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
//[ �T  �v ]�@FMethodList�� BookMark ��ݒ�
//*****************************************************************************
procedure TMethodListForm.SetBookMarkInfo();
var
  i, j: Integer;
  BMLineNo: Integer;
  MethodLineInfo: TMethodLineInfo;
begin
  //��U���ׂĂ��N���A
  for i := 0 to Length(FMethodList) - 1 do
  begin
    FMethodList[i].BookMark := [];
  end;

  //�u�b�N�}�[�N�����[�v
  for i := 0 to 9 do
  begin
    BMLineNo := EditView.BookmarkPos[i].Line;
    if BMLineNo <> 0 then
    begin
      MethodLineInfo := GetMethodLineInfo(BMLineNo, FText);
      if MethodLineInfo.StartRow <> 0 then
      begin
        //���\�b�h�������[�v
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
//[ �T  �v ]�@ListBox��Index�ɊY������TLineInfo���擾
//[ ��  �� ]�@ListBox��Index
//[ �߂�l ]�@TLineInfo
//*****************************************************************************
function TMethodListForm.GetLineInfo(Index: Integer): TLineInfo;
var
  i: Integer;
begin
  i := Integer(ListBox.Items.Objects[Index]);
  Result := FMethodList[i];
end;

//*****************************************************************************
//[ �T  �v ]�@�u������̐ݒ�v���j���[�N���b�N��
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
//[ �T  �v ]�@�u������̍폜�v���j���[�N���b�N��
//            ListBox��DELETE�L�[������
//*****************************************************************************
procedure TMethodListForm.miBookMarkDelClick(Sender: TObject);
begin
  if ListBox.ItemIndex >= 0 then
    ClearBookmark(ListBox.ItemIndex);
end;

//*****************************************************************************
//[ �T  �v ]�@���j���[�u�J�[�\���ʒu�֖߂�v�N���b�N��
//*****************************************************************************
procedure TMethodListForm.miGoBackClick(Sender: TObject);
var
  i: Integer;
begin
  //�t�H�[���N�����̃J�[�\���ʒu��\��
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

  //�ĕ`��
  EditView.Paint;
end;

end.
