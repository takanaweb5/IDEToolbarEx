object MethodListForm: TMethodListForm
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = #12513#12477#12483#12489#19968#35239
  ClientHeight = 359
  ClientWidth = 582
  Color = clBtnFace
  Font.Charset = SHIFTJIS_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = #65325#65331' '#65328#12468#12471#12483#12463
  Font.Style = []
  KeyPreview = True
  Position = poOwnerFormCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  DesignSize = (
    582
    359)
  TextHeight = 13
  object ListBox: TListBox
    Left = 0
    Top = 0
    Width = 582
    Height = 297
    Style = lbOwnerDrawFixed
    Align = alTop
    Anchors = [akLeft, akTop, akRight, akBottom]
    Font.Charset = SHIFTJIS_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = #65325#65331' '#12468#12471#12483#12463
    Font.Style = []
    ItemHeight = 13
    ParentFont = False
    PopupMenu = PopupMenu
    TabOrder = 0
    OnClick = ListBoxClick
    OnDblClick = ListBoxDblClick
    OnDrawItem = ListBoxDrawItem
    OnKeyPress = ListBoxKeyPress
    OnMouseDown = ListBoxMouseDown
    ExplicitLeft = 8
    ExplicitTop = 4
  end
  object cmdOK: TButton
    Left = 489
    Top = 321
    Width = 80
    Height = 24
    Anchors = [akRight, akBottom]
    Caption = #38281#12376#12427
    Default = True
    TabOrder = 5
    OnClick = cmdOKClick
    ExplicitLeft = 487
    ExplicitTop = 313
  end
  object chkMethodName: TCheckBox
    Left = 132
    Top = 335
    Width = 109
    Height = 17
    Anchors = [akLeft, akBottom]
    Caption = #21516#12376#12513#12477#12483#12489#12398#12415
    TabOrder = 4
    ExplicitTop = 327
  end
  object chkSort: TCheckBox
    Left = 15
    Top = 335
    Width = 65
    Height = 17
    Anchors = [akLeft, akBottom]
    Caption = #12477#12540#12488
    TabOrder = 2
    ExplicitTop = 327
  end
  object chkClassName: TCheckBox
    Left = 132
    Top = 309
    Width = 109
    Height = 17
    Anchors = [akLeft, akBottom]
    Caption = #21516#12376#12463#12521#12473#12398#12415
    TabOrder = 3
    ExplicitTop = 301
  end
  object cmbCommand: TComboBox
    Left = 12
    Top = 307
    Width = 109
    Height = 23
    Style = csDropDownList
    Anchors = [akLeft, akBottom]
    Font.Charset = SHIFTJIS_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = #65325#65331' '#12468#12471#12483#12463
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    Items.Strings = (
      #12377#12409#12390
      #12463#12521#12473#23459#35328
      'procedure'
      'function'
      'constructor'
      'destructor')
    ExplicitTop = 299
  end
  object ImageList: TImageList
    Left = 112
    Top = 28
    Bitmap = {
      494C010102000800040010001000FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000400000001000000001002000000000000010
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000004242
      4200424242004242420042424200424242004242420042424200424242000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000084000000840000008400000084000000840000000000000000
      0000000000000000000000000000000000000000000000000000000000004242
      4200000000000000000000FF9C000000000000FF9C0000000000424242004242
      4200000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000084000000840000008400000084000000840000000000000000
      0000000000000000000000000000000000000000000000000000000000004242
      42000000000000FF9C000000000000FF9C000000000000FF9C00424242004242
      4200000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000084000000840000008400000084000000840000000000000000
      0000000000000000000000000000000000000000000000000000000000004242
      4200000000000000000000FF9C000000000000FF9C0000000000424242004242
      4200000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000084000000840000008400000084000000840000000000000000
      0000000000000000000000000000000000000000000000000000000000004242
      42000000000000FF9C000000000000FF9C000000000000FF9C00424242004242
      4200000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000084000000840000008400000084000000840000000000000000
      0000000000000000000000000000000000000000000000000000000000004242
      4200000000000000000000FF9C000000000000FF9C0000000000424242004242
      4200000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000084000000840000008400000084000000840000000000000000
      0000000000000000000000000000000000000000000000000000000000004242
      42000000000000FF9C000000000000FF9C000000000000FF9C00424242004242
      4200000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000004242
      4200424242004242420042424200424242004242420042424200424242004242
      4200000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000042424200424242004242420042424200424242004242
      4200000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000424D3E000000000000003E000000
      2800000040000000100000000100010000000000800000000000000000000000
      000000000000000000000000FFFFFF00FFFFFFFF00000000FFFFFFFF00000000
      FFFFFFFF00000000FFFFFFFF00000000E01FE01F00000000E80FED4F00000000
      E80FEA8F00000000E80FED4F00000000E80FEA8F00000000E80FED4F00000000
      E80FEA8F00000000E00FE00F00000000FC0FFC0F00000000FFFFFFFF00000000
      FFFFFFFF00000000FFFFFFFF0000000000000000000000000000000000000000
      000000000000}
  end
  object ApplicationEvents: TApplicationEvents
    Left = 56
    Top = 32
  end
  object PopupMenu: TPopupMenu
    Left = 40
    Top = 80
    object miBookMarkSet: TMenuItem
      Caption = #12375#12362#12426#12398#35373#23450
      OnClick = miBookMarkSetClick
    end
    object miBookMarkDel: TMenuItem
      Caption = #12375#12362#12426#12398#21066#38500
      ShortCut = 46
      OnClick = miBookMarkDelClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object miGoBack: TMenuItem
      Caption = #32232#38598#34892#12408#25147#12427'(&I)'
      OnClick = miGoBackClick
    end
  end
end
