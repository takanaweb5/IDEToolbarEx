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
//[イベント]　フォーム作成時
//[ 概  要 ]　初期処理
//*****************************************************************************
procedure TFindDlg.FormCreate(Sender: TObject);
begin
  FStringList := TStringList.Create;
end;

//*****************************************************************************
//[イベント]　フォーム廃棄時
//*****************************************************************************
procedure TFindDlg.FormDestroy(Sender: TObject);
begin
  FStringList.Free;
end;

//*****************************************************************************
//[イベント]　フォーム表示時
//[ 概  要 ]　画面の初期設定
//*****************************************************************************
procedure TFindDlg.FormShow(Sender: TObject);
begin
  cmbFindWord.Items.Assign(FStringList);
  cmbFindWord.Text := FFindText;
  cmdOK.Enabled := (cmbFindWord.Text <> '');
end;

//*****************************************************************************
//[イベント]　検索文字列変更時
//[ 概  要 ]　ＯＫボタンのEnabled設定
//*****************************************************************************
procedure TFindDlg.cmbFindWordChange(Sender: TObject);
begin
  cmdOK.Enabled := (cmbFindWord.Text <> '');
end;

//*****************************************************************************
//[イベント]　ＯＫボタンクリック時
//[ 概  要 ]　検索文字列を次回検索時に選択出来るようにStringListに保存
//*****************************************************************************
procedure TFindDlg.cmdOKClick(Sender: TObject);
var
  FindIndex: Integer;
begin
  //2重登録のチェック
  FindIndex := FStringList.IndexOf(FindText);
  if FindIndex >= 0 then
  begin
    //登録済みを１番目に移動
    FStringList.Move(FindIndex, 0);
  end
  else
  begin
    //StringListに登録
    FStringList.Insert(0, FindText);
  end;

  ModalResult := mrOk;
  Hide;
end;

//*****************************************************************************
//[イベント]　キャンセルボタンクリック時
//[ 概  要 ]　キャンセル
//*****************************************************************************
procedure TFindDlg.cmdCancelClick(Sender: TObject);
begin
  ModalResult := mrCancel;
  Hide;
end;

//*****************************************************************************
//[プロパティ]　FindText
//[ 概  要 ]　検索文字列を設定する
//*****************************************************************************
procedure TFindDlg.SetFindText(const Value: string);
begin
  FFindText := Value;
end;

//*****************************************************************************
//[プロパティ]　FindText
//[ 概  要 ]　検索文字列を取得する
//*****************************************************************************
function TFindDlg.GetFindText: string;
begin
  Result := cmbFindWord.Text;
end;

end.
