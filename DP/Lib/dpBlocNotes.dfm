inherited FBlocNotes: TFBlocNotes
  Left = 391
  Width = 471
  Height = 400
  Caption = 'Commentaire'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 331
    Width = 463
    inherited PBouton: TToolWindow97
      ClientWidth = 463
      ClientAreaWidth = 463
      inherited BValider: TToolbarButton97
        Left = 367
      end
      inherited BFerme: TToolbarButton97
        Left = 399
      end
      inherited HelpBtn: TToolbarButton97
        Left = 431
      end
      inherited BImprimer: TToolbarButton97
        Left = 335
      end
    end
  end
  object OleBlocNotes: THRichEditOLE [1]
    Left = 0
    Top = 0
    Width = 463
    Height = 331
    Align = alClient
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    OnChange = OleBlocNotesChange
    Margins.Top = 0
    Margins.Bottom = 0
    Margins.Left = 0
    Margins.Right = 0
    ContainerName = 'Document'
    ObjectMenuPrefix = '&Object'
    LinesRTF.Strings = (
      
        '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fnil Aria' +
        'l;}}'
      
        '{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard\f0\fs16' +
        ' '
      '\par '
      '\par }')
  end
end
