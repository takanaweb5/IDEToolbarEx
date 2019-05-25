unit FindForm;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms, StdCtrls;

type
  TFindDlg = class(TForm)
    Label1: TLabel;
    cmbFindWord: TComboBox;
    chkCase: TCheckBox;
    chkWord: TCheckBox;
    chkRegEx: TCheckBox;
    cmdOK: TButton;
    cmdCancel: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure cmbFindWordChange(Sender: TObject);
    procedure cmdCancelClick(Sender: TObject);
    procedure cmdOKClick(Sender: TObject);
  private
    FFindText: string;
    FStringList: TStringList;
    procedure SetFindText(const Value: string);
    function GetFindText: string;
  public
    property FindText: string read GetFindText write SetFindText;
  end;

var
  FindDlg: TFindDlg;

implementation
{$R *.dfm}

uses
  Dialogs;

//*****************************************************************************
//[�C�x���g]�@�t�H�[���쐬��
//[ �T  �v ]�@��������
//*****************************************************************************
procedure TFindDlg.FormCreate(Sender: TObject);
begin
  FStringList := TStringList.Create;
end;

//*****************************************************************************
//[�C�x���g]�@�t�H�[���p����
//*****************************************************************************
procedure TFindDlg.FormDestroy(Sender: TObject);
begin
  FStringList.Free;
end;

//*****************************************************************************
//[�C�x���g]�@�t�H�[���\����
//[ �T  �v ]�@��ʂ̏����ݒ�
//*****************************************************************************
procedure TFindDlg.FormShow(Sender: TObject);
begin
  cmbFindWord.Items.Assign(FStringList);
  cmbFindWord.Text := FFindText;
  cmdOK.Enabled := (cmbFindWord.Text <> '');
end;

//*****************************************************************************
//[�C�x���g]�@����������ύX��
//[ �T  �v ]�@�n�j�{�^����Enabled�ݒ�
//*****************************************************************************
procedure TFindDlg.cmbFindWordChange(Sender: TObject);
begin
  cmdOK.Enabled := (cmbFindWord.Text <> '');
end;

//*****************************************************************************
//[�C�x���g]�@�n�j�{�^���N���b�N��
//[ �T  �v ]�@��������������񌟍����ɑI���o����悤��StringList�ɕۑ�
//*****************************************************************************
procedure TFindDlg.cmdOKClick(Sender: TObject);
var
  FindIndex: Integer;
begin
  //2�d�o�^�̃`�F�b�N
  FindIndex := FStringList.IndexOf(FindText);
  if FindIndex >= 0 then
  begin
    //�o�^�ς݂��P�ԖڂɈړ�
    FStringList.Move(FindIndex, 0);
  end
  else
  begin
    //StringList�ɓo�^
    FStringList.Insert(0, FindText);
  end;

  ModalResult := mrOk;
  Hide;
end;

//*****************************************************************************
//[�C�x���g]�@�L�����Z���{�^���N���b�N��
//[ �T  �v ]�@�L�����Z��
//*****************************************************************************
procedure TFindDlg.cmdCancelClick(Sender: TObject);
begin
  ModalResult := mrCancel;
  Hide;
end;

//*****************************************************************************
//[�v���p�e�B]�@FindText
//[ �T  �v ]�@�����������ݒ肷��
//*****************************************************************************
procedure TFindDlg.SetFindText(const Value: string);
begin
  FFindText := Value;
end;

//*****************************************************************************
//[�v���p�e�B]�@FindText
//[ �T  �v ]�@������������擾����
//*****************************************************************************
function TFindDlg.GetFindText: string;
begin
  Result := cmbFindWord.Text;
end;

end.
