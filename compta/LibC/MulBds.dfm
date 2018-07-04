inherited FMulBds: TFMulBds
  Left = 231
  Top = 234
  Caption = 'Modification Balances de situation'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: THPageControl2
    Height = 79
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Height = 51
      end
      object HLEXERCICE: THLabel
        Left = 9
        Top = 17
        Width = 41
        Height = 13
        Caption = 'Exercice'
        FocusControl = HB_EXERCICE
      end
      object HLDATE1: THLabel
        Left = 258
        Top = 17
        Width = 29
        Height = 13
        Caption = 'D'#233'but'
        FocusControl = DATE1
      end
      object HLFIN: THLabel
        Left = 405
        Top = 17
        Width = 14
        Height = 13
        Caption = 'Fin'
        FocusControl = DATE2
      end
      object HB_EXERCICE: THValComboBox
        Left = 55
        Top = 13
        Width = 193
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 0
        OnChange = HB_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object HB_DATE1: THCritMaskEdit
        Left = 141
        Top = 27
        Width = 73
        Height = 21
        Color = clYellow
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 1
        Text = '  /  /    '
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = od1900
        ControlerDate = True
      end
      object HB_DATE2: THCritMaskEdit
        Left = 235
        Top = 27
        Width = 69
        Height = 21
        Color = clYellow
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 2
        Text = '  /  /    '
        Visible = False
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        DefaultDate = od2099
        ControlerDate = True
      end
      object XX_WHERE: TEdit
        Left = 336
        Top = 27
        Width = 111
        Height = 21
        Color = clYellow
        TabOrder = 3
        Text = 'HB_TYPEBAL="BDS"'
        Visible = False
      end
      object DATE1: THValComboBox
        Left = 292
        Top = 13
        Width = 105
        Height = 21
        Style = csDropDownList
        Enabled = False
        ItemHeight = 0
        TabOrder = 4
        OnChange = DATE1Change
        TagDispatch = 0
      end
      object DATE2: THValComboBox
        Left = 424
        Top = 13
        Width = 105
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 5
        OnChange = DATE2Change
        TagDispatch = 0
      end
    end
    inherited PComplement: THTabSheet
      TabVisible = False
      inherited Bevel2: TBevel
        Height = 51
      end
    end
    inherited PAvance: THTabSheet
      TabVisible = False
      inherited Bevel4: TBevel
        Height = 51
      end
    end
    inherited PSQL: THTabSheet
      inherited Bevel3: TBevel
        Height = 51
      end
      inherited Z_SQL: THSQLMemo
        Height = 51
      end
    end
  end
  inherited Dock971: TDock97
    Top = 79
  end
  inherited FListe: THDBGrid
    Top = 120
    Height = 198
    MultiSelection = True
    MultiFieds = 'HB_TYPEBAL;HB_EXERCICE;HB_DATE1;HB_DATE2;'
  end
  inherited Dock: TDock97
    inherited PanelBouton: TToolWindow97
      inherited bSelectAll: TToolbarButton97
        Visible = True
      end
      inherited Binsert: TToolbarButton97
        Visible = True
        OnClick = BinsertClick
      end
      object BDelete: TToolbarButton97
        Left = 243
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Supprimer'
        DisplayMode = dmGlyphOnly
        Caption = 'Supprimer'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BDeleteClick
        GlobalIndexImage = 'Z0005_S16G1'
      end
    end
  end
  inherited PanVBar: THPanel
    Top = 120
    Height = 198
  end
  inherited Q: THQuery
    Distinct = True
  end
  object HMulBds: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;Aucune balance s'#233'lectionn'#233'e;W;O;O;O;'
      '1;?caption?;Confirmez-vous la suppression ?;Q;YNC;N;C;'
      '2;?caption?;Impossible de supprimer la balance;E;O;O;O;')
    Left = 394
    Top = 242
  end
end
