object FindDlg: TFindDlg
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = #26908#32034
  ClientHeight = 122
  ClientWidth = 481
  Color = clBtnFace
  Font.Charset = SHIFTJIS_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = #65325#65331' '#65328#12468#12471#12483#12463
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  DesignSize = (
    481
    122)
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 10
    Width = 84
    Height = 13
    Caption = #26908#32034#25991#23383#21015'(&T):'
    FocusControl = cmbFindWord
  end
  object cmbFindWord: TComboBox
    Left = 104
    Top = 8
    Width = 369
    Height = 21
    Anchors = [akLeft, akTop, akRight]
    Font.Charset = SHIFTJIS_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #65325#65331' '#12468#12471#12483#12463
    Font.Style = []
    ItemHeight = 13
    ParentFont = False
    TabOrder = 0
    OnChange = cmbFindWordChange
  end
  object chkCase: TCheckBox
    Left = 44
    Top = 40
    Width = 141
    Height = 25
    Caption = #22823'/'#23567#25991#23383#12398#21306#21029'(&C)'
    TabOrder = 1
  end
  object chkWord: TCheckBox
    Left = 44
    Top = 64
    Width = 141
    Height = 25
    Caption = #12527#12540#12489#26908#32034'(&W)'
    TabOrder = 2
  end
  object chkRegEx: TCheckBox
    Left = 44
    Top = 88
    Width = 141
    Height = 25
    Caption = #27491#35215#34920#29694#26908#32034'(&R)'
    TabOrder = 3
  end
  object cmdOK: TButton
    Left = 203
    Top = 82
    Width = 121
    Height = 26
    Anchors = [akRight, akBottom]
    Caption = 'OK'
    Default = True
    TabOrder = 4
    OnClick = cmdOKClick
  end
  object cmdCancel: TButton
    Left = 343
    Top = 82
    Width = 121
    Height = 26
    Anchors = [akRight, akBottom]
    Cancel = True
    Caption = #12461#12515#12531#12475#12523
    TabOrder = 5
    OnClick = cmdCancelClick
  end
end
