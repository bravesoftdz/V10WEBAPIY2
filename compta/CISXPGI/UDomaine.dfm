inherited FFicheDomaine: TFFicheDomaine
  Left = 221
  Top = 254
  Caption = 'Domaines'
  PixelsPerInch = 96
  TextHeight = 13
  inherited FListe: THDBGrid
    DataSource = dmImport.DataSource3
    Options = [dgEditing, dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs]
  end
  inherited DBNav: TDBNavigator
    Hints.Strings = ()
  end
  inherited DS: TDataSource
    DataSet = nil
  end
  inherited T: THTable
    DatabaseName = 'GLOBAL'
    MasterSource = dmImport.DataSource3
  end
end
