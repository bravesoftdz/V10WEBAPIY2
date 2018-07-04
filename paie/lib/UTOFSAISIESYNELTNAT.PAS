{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 14/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOFSAISIESYNELTNAT ()
Mots clefs ... : TOF;UTOFSAISIESYNELTNAT
*****************************************************************
PT1  21/06/2007 FC V_72 FQ 14431 Doublon dans la liste des valeurs liées à une table dynamique
PT2  21/09/2007 FC V_80 FQ 14800 Filtrer les valeurs d'un élément dynamique en fonction de la convention collective
PT3  08/11/2007 FC V_80 FQ 14922 CWAS Erreur SQL quand click sur cellule pour saisir un élément dynamique
PT4  16/11/2007 FC V_80 FQ 14931 Pb filtrage des valeurs d'un élément quand plusieurs prédéfinis paramétrés
PT5  21/01/2008 FC V_81 FQ 15144 Cas particulier de la saisie d'une valeur boite à cocher
PT6  26/02/2008 FC V_81 CWAS : violation d'accès quand séparateur décimale paramètres régionaux est le '.'
PT7  17/04/2008 GGU V81 FQ 15361 Factorisation du code en vue de la modification
}
Unit UTOFSAISIESYNELTNAT ;

Interface

Uses StdCtrls,
     Controls,
     Graphics,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     forms,
     uTob,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     P5Util,
     Grids,
     HTB97,
     LookUp;

Const
  ColStatus        = 1;
  ColCodeElt       = 2;
  ColLibelle       = 3;
  ColTypNiv        = 4;
  ColLIBVALNIV     = 5;
  ColMontantEuro   = 6;
//  ColMontant       = 7;
  ColConvention    = 7;
  ColDatVal        = 8;
  ColTheme         = 9;
  ColAbrege        = 10;
  ColMonetaire     = 11;
  ColBlocNote      = 12;
  ColDecalMois     = 13;
  ColRegimeAlsace  = 14;
  ColETABLISSEMENT = 15;
  ColSALARIE       = 16;
  ColCODEPOP       = 17;
  ColValNiv        = 18;
  ColPredefini     = 19;
  ColNoDossier     = 20;

  ColCLStatus        = 1;
  ColCLLibelle       = 2;
  ColCLCodLibre      = 3;
  ColCLSALARIE       = 4;
  ColCLLibelleVal    = 5;
  ColCLValeur        = 6;
  ColCLDatVal        = 7;
  ColCLTheme         = 8;
  ColCLAbrege        = 9;
  ColCLETABLISSEMENT = 10;
  ColCLTypNiv        = 11;
  ColCLCODEPOP       = 12;
  ColCLValNiv        = 13;
  ColCLPredefini     = 14;
  ColCLNoDossier     = 15;

  Type
  TOF_UTOFSAISIESYNELTNAT = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    GEltNat, GCodLibre : THGrid;
    StChampGrid, TypeElipsis, StChampGridCL : String;
    TobEltNat, TobCodLibre : Tob;
    ClePGSynEltNAt : String;
    AucunEltNational, AucunEltDynamique : Boolean;
    procedure MiseEnFormeGrille;
    procedure ChargerDonnees;
    procedure GrillePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrilleColEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GrilleElipsisClick(Sender: TObject);
    procedure GCodLibrePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure GCodLibreCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GCodLibreCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure ZoneSuivanteOuOk(Grille: THGrid; ACol, ARow: Longint; var Cancel: boolean);
    function ZoneAccessible(CodLibre: String; var ACol, ARow: Longint): Boolean;
    procedure ValidationSaisie(Sender : TObject);
    procedure AnnulerModifs(Sender : TObject);
  end ;

Implementation

uses TntGrids,
  PGTablesDyna; //PT7

procedure TOF_UTOFSAISIESYNELTNAT.OnArgument (S : String ) ;
var
  Titres : HTStringList;
  BValider,BDefaire : TToolBarButton97;
begin
  Inherited ;

  ClePGSynEltNAt := S;

  GEltNat := THGrid(GeTControl('GELTNAT'));
  GEltNat.DBIndicator := True;
  GEltNat.ColCount := 21;
  Titres := HTStringList.Create;
  Titres.Insert(0,'');
  Titres.Insert(ColStatus        ,'');
  Titres.Insert(ColCodeElt       ,TraduireMemoire('Elt nat'));
  Titres.Insert(ColLibelle       ,TraduireMemoire('Libellé'));
  Titres.Insert(ColTypNiv        ,TraduireMemoire('Niveau'));
  Titres.Insert(ColLIBVALNIV     ,TraduireMemoire('Etab/Pop/Sal'));
  Titres.Insert(ColMontantEuro   ,TraduireMemoire('Valeur'));
//  Titres.Insert(ColMontant       ,TraduireMemoire('Valeur Alsace'));
  Titres.Insert(ColConvention    ,TraduireMemoire('Conv.'));
  Titres.Insert(ColDatVal        ,TraduireMemoire('D. validité'));
  Titres.Insert(ColTheme         ,TraduireMemoire('Thème'));
  Titres.Insert(ColAbrege        ,TraduireMemoire('Abrégé'));
  Titres.Insert(ColMonetaire     ,TraduireMemoire('Monétaire'));
  Titres.Insert(ColBlocNote      ,TraduireMemoire('Bloc notes'));
  Titres.Insert(ColDecalMois     ,TraduireMemoire('Décal. mois'));
  Titres.Insert(ColRegimeAlsace  ,TraduireMemoire('R. Alsace'));
  Titres.Insert(ColETABLISSEMENT ,'');
  Titres.Insert(ColSALARIE       ,'');
  Titres.Insert(ColCODEPOP       ,'');
  Titres.Insert(ColValNiv        ,'');
  Titres.Insert(ColPredefini     ,'');
  Titres.Insert(ColNoDossier     ,'');
  GEltNat.Titres := Titres;
  Titres.free;
  if (GEltNat <> nil) then
  begin
    GEltNat.PostDrawCell := GrillePostDrawCell;
    GEltNat.OnCellExit   := GrilleCellExit;
    GEltNat.OnCellEnter  := GrilleCellEnter;
  end;

  GCodLibre := THGrid(GeTControl('GCODLIBRE'));
  GCodLibre.DBIndicator := True;
  GCodLibre.ColCount := 16;
  Titres := HTStringList.Create;
  Titres.Insert(0,'');
  Titres.Insert(ColCLStatus        ,'');
  Titres.Insert(ColCLLibelle       ,TraduireMemoire('Libellé'));
  Titres.Insert(ColCLCodLibre      ,TraduireMemoire('Code table'));
  Titres.Insert(ColCLSALARIE         ,TraduireMemoire('Salarié'));
  Titres.Insert(ColCLLibelleVal    ,TraduireMemoire('Elément dynamique'));
  Titres.Insert(ColCLValeur        ,TraduireMemoire('Libellé de l''élément'));
  Titres.Insert(ColCLDatVal        ,TraduireMemoire('Date validité'));
  Titres.Insert(ColCLTheme         ,'');
  Titres.Insert(ColCLAbrege        ,'');
  Titres.Insert(ColCLETABLISSEMENT ,'');
  Titres.Insert(ColCLTypNiv        ,'');
  Titres.Insert(ColCLCODEPOP       ,'');
  Titres.Insert(ColCLValNiv        ,'');
  Titres.Insert(ColCLPredefini     ,'');
  Titres.Insert(ColCLNoDossier     ,'');
  GCodLibre.Titres := Titres;
  Titres.free;
  if (GCodLibre <> nil) then
  begin
    GCodLibre.OnColEnter   := GrilleColEnter;
    GCodLibre.OnElipsisClick := GrilleElipsisClick;
    GCodLibre.PostDrawCell := GCodLibrePostDrawCell;
    GCodLibre.OnCellExit   := GCodLibreCellExit;
    GCodLibre.OnCellEnter  := GCodLibreCellEnter;
  end;

  BValider := TToolBarButton97(GetControl('VALIDSAISIE'));
  if BValider <> nil then
    BValider.OnClick := ValidationSaisie;

  BDefaire := TToolBarButton97(GetControl('BDEFAIRE'));
  if BDefaire <> nil then
    BDefaire.OnClick := AnnulerModifs;

  MiseEnFormeGrille;
  ChargerDonnees;
end ;

Procedure TOF_UTOFSAISIESYNELTNAT.AnnulerModifs(Sender : TObject);
var
  i : Integer;
begin
  TobEltNat.PutGridDetail(GEltNat,False,False,StChampGrid,False);
  GEltNat.SortedCol := -1;
  for i := 1 to (GEltNat.RowCount - 1) do
    GEltNat.CellValues[ColStatus,i] := '#ICO#43';

  TobCodLibre.PutGridDetail(GCodLibre,False,False,StChampGridCL,False);
  GCodLibre.SortedCol := -1;
  for i := 1 to (GCodLibre.RowCount - 1) do
  begin
    GCodLibre.CellValues[ColCLStatus,i] := '#ICO#43';
    GCodLibre.CellValues[ColCLLibelleVal,i] := '';
    GCodLibre.CellValues[ColCLValeur,i] := '';
  end;
end;

procedure TOF_UTOFSAISIESYNELTNAT.OnClose ;
begin
  Inherited ;
  FreeAndNil(TobEltNat);
  FreeAndNil(TobCodLibre);
end ;

procedure TOF_UTOFSAISIESYNELTNAT.MiseEnFormeGrille;
begin
  // Mise en forme des colonnes
  GEltNAt.ColEditables[ColStatus]        := False;
  GEltNAt.ColWidths[ColStatus]           := 20;

  GEltNAt.ColEditables[ColCodeElt]       := False;
  GEltNAt.ColWidths[ColCodeElt]          := 40;

  GEltNAt.ColEditables[ColLibelle]       := False;
  GEltNAt.ColWidths[ColLibelle]          := 105;

  GEltNAt.ColEditables[ColTypNiv]        := False;
  GEltNAt.ColFormats[ColTypNiv]          := 'CB=PGNIVEAUAVDOS';
  GEltNAt.ColWidths[ColTypNiv]           := 80;

  GEltNAt.ColEditables[ColLibValNiv]     := False;
  GEltNAt.ColWidths[ColLibValNiv]        := 80;

  GEltNAt.ColEditables[ColMontantEuro]   := True;
  GEltNAt.ColTypes[ColMontantEuro]       := 'R';
  GEltNAt.ColFormats[ColMontantEuro]     := '# ##0.00';
  GEltNAt.ColAligns[ColMontantEuro]      := taLeftJustify;
  GEltNAt.ColWidths[ColMontantEuro]      := 60;

{  GEltNAt.ColEditables[ColMontant]       := True;
  GEltNAt.ColTypes[ColMontant]           := 'R';
  GEltNAt.ColFormats[ColMontant]         := '# ##0.00';
  GEltNAt.ColAligns[ColMontant]          := taLeftJustify;
  GEltNAt.ColWidths[ColMontant]          := 60;}

  GEltNAt.ColEditables[ColConvention]    := True;
  GEltNAt.ColFormats[ColConvention]      := 'CB=PGCONVENTIONS';
  GEltNAt.ColAligns[ColConvention]       := taLeftJustify;
  GEltNAt.ColWidths[ColConvention]       := 60;

  GEltNAt.ColEditables[ColDatVal]        := True;
  GEltNAt.ColTypes[ColDatVal]            := 'D';
  GEltNAt.ColFormats[ColDatVal]          := ShortDateFormat;
  GEltNAt.ColAligns[ColDatVal]           := taCenter;
  GEltNAt.ColWidths[ColDatVal]           := 65;

  GEltNAt.ColEditables[ColTheme]         := True;
  GEltNAt.ColFormats[ColTheme]           := 'CB=PGTHEMEELEMENTNAT';
  GEltNAt.ColAligns[ColTheme]            := taLeftJustify;
  GEltNAt.ColWidths[ColTheme]            := 64;

  GEltNAt.ColEditables[ColAbrege]        := False;
  GEltNAt.ColWidths[ColAbrege]           := -1;

  GEltNAt.ColEditables[ColMonetaire]     := False;
  GEltNAt.ColTypes[ColMonetaire]         := 'B';
  GEltNAt.ColFormats[ColMonetaire]       := IntToStr(Ord(csCheckBox));
  GEltNAt.ColAligns[ColMonetaire]        := taCenter;
  GEltNAt.ColWidths[ColMonetaire]        := 65;

  GEltNAt.ColEditables[ColBlocNote]      := False;
  GEltNAt.ColWidths[ColBlocNote]         := -1;

  GEltNAt.ColEditables[ColDecalMois]     := False;
  GEltNAt.ColTypes[ColDecalMois]         := 'B';
  GEltNAt.ColFormats[ColDecalMois]       := IntToStr(Ord(csCheckBox));
  GEltNAt.ColAligns[ColDecalMois]        := taCenter;
  GEltNAt.ColWidths[ColDecalMois]        := 80;

  GEltNAt.ColEditables[ColRegimeAlsace]  := False;
  GEltNAt.ColTypes[ColRegimeAlsace]      := 'B';
  GEltNAt.ColFormats[ColRegimeAlsace]    := IntToStr(Ord(csCheckBox));
  GEltNAt.ColAligns[ColRegimeAlsace]     := taCenter;
  GEltNAt.ColWidths[ColRegimeAlsace]     := 65;

  GEltNAt.ColEditables[ColETABLISSEMENT] := False;
  GEltNAt.ColWidths[ColETABLISSEMENT]    := -1;
  GEltNAt.ColEditables[ColSALARIE]       := False;
  GEltNAt.ColWidths[ColSALARIE]          := -1;
  GEltNAt.ColEditables[ColCODEPOP]       := False;
  GEltNAt.ColWidths[ColCODEPOP]          := -1;
  GEltNAt.ColEditables[ColVALNIV]        := False;
  GEltNAt.ColWidths[ColVALNIV]           := -1;
  GEltNAt.ColEditables[ColPredefini]     := False;
  GEltNAt.ColWidths[ColPredefini]        := -1;
  GEltNAt.ColEditables[ColNoDossier]     := False;
  GEltNAt.ColWidths[ColNoDossier]        := -1;


  GCodLibre.ColEditables[ColCLStatus]        := False;
  GCodLibre.ColWidths[ColCLStatus]           := 20;

  GCodLibre.ColEditables[ColCLCodLibre]      := False;
  GCodLibre.ColWidths[ColCLCodLibre]         := 100;

  GCodLibre.ColEditables[ColCLLibelle]       := False;
  GCodLibre.ColWidths[ColCLLibelle]          := 170;

  GCodLibre.ColEditables[ColCLSALARIE]        := False;
  GCodLibre.ColWidths[ColCLSALARIE]           := 120;

  GCodLibre.ColEditables[ColCLLibelleVal]    := True;
  GCodLibre.ColWidths[ColCLLibelleVal]       := 140;

  GCodLibre.ColEditables[ColCLValeur]        := True;
  GCodLibre.ColWidths[ColCLValeur]           := 135;

  GCodLibre.ColEditables[ColCLDatVal]        := True;
  GCodLibre.ColTypes[ColCLDatVal]            := 'D';
  GCodLibre.ColFormats[ColCLDatVal]          := ShortDateFormat;
  GCodLibre.ColWidths[ColCLDatVal]           := 100;
  GCodLibre.ColAligns[ColCLDatVal]           := taCenter;

  GCodLibre.ColEditables[ColCLTheme]         := False;
  GCodLibre.ColWidths[ColCLTheme]            := -1;
  GCodLibre.ColEditables[ColCLAbrege]        := False;
  GCodLibre.ColWidths[ColCLAbrege]           := -1;
  GCodLibre.ColEditables[ColCLETABLISSEMENT] := False;
  GCodLibre.ColWidths[ColCLETABLISSEMENT]    := -1;
  GCodLibre.ColEditables[ColCLTypNiv]       := False;
  GCodLibre.ColWidths[ColCLTypNiv]          := -1;
  GCodLibre.ColEditables[ColCLCODEPOP]       := False;
  GCodLibre.ColWidths[ColCLCODEPOP]          := -1;
  GCodLibre.ColEditables[ColCLVALNIV]        := False;
  GCodLibre.ColWidths[ColCLVALNIV]           := -1;
  GCodLibre.ColEditables[ColCLPredefini]     := False;
  GCodLibre.ColWidths[ColCLPredefini]        := -1;
  GCodLibre.ColEditables[ColCLNoDossier]     := False;
  GCodLibre.ColWidths[ColCLNoDossier]        := -1;
end;

procedure TOF_UTOFSAISIESYNELTNAT.ChargerDonnees;
var
  Q:TQuery;
  i : Integer;
  TobCL, TobEL : Tob;
begin
  if TobEltNat = nil then FreeAndNil(TobEltNat);
  if TobCodLibre = nil then FreeAndNil(TobCodLibre);

  // Chargement des éléments manquants : PEY_TYPUTI = "1"
  Q := OpenSQL('SELECT PEY_CODEELT,PEY_LIBELLE,PEY_NODOSSIER,PEY_TYPNIV,PEY_VALNIV,PEY_DATVAL,PEY_MONTANTEURO,'
    + 'PEY_PREDEFINI,PEY_CONVENTION,PEY_THEMEELT,PEY_ABREGE,PEY_MONETAIRE,PEY_BLOCNOTE,'
    + 'PEY_DECALMOIS,PEY_REGIMEALSACE,PEY_ETABLISSEMENT,PEY_SALARIE,PEY_CODEPOP,PEY_LIBVALNIV'
    + ' FROM PGSYNELTNAT WHERE PEY_TYPUTI = "1" AND PEY_NODOSSIER = "' + ClePGSynEltNAt + '"'
    + ' ORDER BY PEY_CODEELT,PEY_PREDEFINI,PEY_TYPNIV,PEY_VALNIV,PEY_DATVAL',True);
  TobEltNat := Tob.Create('SyntheseEltNat',Nil,-1);
  TobEltNat.LoadDetailDB('PGSYNELTNAT','','',Q,False);
  ferme(Q);

  StChampGrid := 'STATUS;PEY_CODEELT;PEY_LIBELLE;PEY_TYPNIV;PEY_LIBVALNIV;PEY_MONTANTEURO;'
    + 'PEY_CONVENTION;PEY_DATVAL;PEY_THEMEELT;PEY_ABREGE;PEY_MONETAIRE;PEY_BLOCNOTE;PEY_DECALMOIS;'
    + 'PEY_REGIMEALSACE;PEY_ETABLISSEMENT;PEY_SALARIE;PEY_CODEPOP;PEY_VALNIV;PEY_PREDEFINI;PEY_NODOSSIER';

  AucunEltNational := False;
  if TobEltNat.Detail.Count = 0 then
  begin
    TobEL := TOB.Create('PGSYNELTNAT', TobEltNat, -1);
    TobEL.PutValue('PEY_CODEELT','Aucun élément à saisir');
    TobEL.PutValue('STATUS','');
    AucunEltNational := True;
  end;
  GEltNat.RowCount := TobEltNat.Detail.Count + 1;
  TobEltNat.PutGridDetail(GEltNat,False,False,StChampGrid,False);

  // ICO27 Smiley rouge; ICO28 Smiley jaune; ICO35 grand drapeau rouge; ICO37 grand drapeau vert
  // ICO43 petit drapeau rouge; ICO44 petit drapeau vert
  //  GEltNat.CellValues[ColStatus,i] := '#ICO#27';
  if not AucunEltNational then
  begin
    for i := 1 to (GEltNat.RowCount - 1) do
      GEltNat.CellValues[ColStatus,i] := '#ICO#43';

    GEltNat.Row := 1;
    GEltNat.Col := 1;
    GEltNat.SetFocus;
  end
  else
    GEltNat.Enabled := False;

  // Chargement des codes libres manquants : PEY_TYPUTI = "3"
  Q := OpenSQL('SELECT PEY_LIBELLE,PEY_CODEELT,PEY_TYPNIV,PEY_DATVAL,PEY_MONTANTEURO,'
    + 'PEY_PREDEFINI,PEY_THEMEELT,PEY_ABREGE,'
    + 'PEY_ETABLISSEMENT,PEY_SALARIE,PEY_CODEPOP,PEY_LIBVALNIV,PEY_VALNIV,PEY_NODOSSIER'
    + ' FROM PGSYNELTNAT WHERE PEY_TYPUTI = "3" AND PEY_NODOSSIER = "' + ClePGSynEltNAt + '"'
    + ' ORDER BY PEY_CODEELT,PEY_PREDEFINI,PEY_TYPNIV,PEY_VALNIV,PEY_DATVAL',True);
  TobCodLibre := Tob.Create('SyntheseEltNat',Nil,-1);
  TobCodLibre.LoadDetailDB('PGSYNELTNAT','','',Q,False);
  ferme(Q);

  StChampGridCL := 'STATUS;PEY_LIBELLE;PEY_CODEELT;PEY_SALARIE;LIBVAL;VALEUR;PEY_DATVAL;PEY_THEMEELT;'
    + 'PEY_ABREGE;PEY_ETABLISSEMENT;PEY_TYPNIV;PEY_CODEPOP;PEY_VALNIV;PEY_PREDEFINI;PEY_NODOSSIER';

  AucunEltDynamique := False;
  if TobCodLibre.Detail.Count = 0 then
  begin
    TobCL := TOB.Create('PGSYNELTNAT', TobCodLibre, -1);
    TobCL.PutValue('PEY_LIBELLE','Aucun élément à saisir');
    TobCL.PutValue('STATUS','');
{    TobEL.PutValue('PEY_TYPUTI','3');
    TobEL.PutValue('PEY_NODOSSIER',ClePGSynEltNAt);
    TobCL.InsertOrUpdateDB;}
    AucunEltDynamique := True;
  end;
  GCodLibre.RowCount := TobCodLibre.Detail.Count + 1;
  TobCodLibre.PutGridDetail(GCodLibre,False,False,StChampGridCL,False);

  if not AucunEltDynamique then
  begin
    for i := 1 to (GCodLibre.RowCount - 1) do
      GCodLibre.CellValues[ColCLStatus,i] := '#ICO#43';

    GCodLibre.Row := 1;
    GCodLibre.Col := 1;
    GCodLibre.SetFocus;
  end
  else
    GCodLibre.Enabled := False;
end;

procedure TOF_UTOFSAISIESYNELTNAT.GrillePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
//var
//  Q : TQuery;
begin
  if ARow <> 0 then
  begin
    If   (ACol = ColCodeElt) or (ACol = ColLibelle) or (ACol = ColPredefini) or (ACol = ColConvention)
      or (ACol = ColTypNiv) or (ACol = ColLibValNiv) or (ACol = ColTheme) or (ACol = ColAbrege)
      or (ACol = ColMonetaire) or (ACol = ColBlocNote) or (ACol = ColDecalMois) or (ACol = ColRegimeAlsace)
      or (ACol = ColStatus) or AucunEltNational then
        GridGriseCell(GEltNat, Acol, Arow, Canvas);
  end
  else
    GEltNat.Enabled := False;
{  if (GEltNat.CellValues[ColPredefini,ARow] = '') and (ACol = ColMontantEuro) then // or (ACol = ColMontant))
  begin
    if (GEltNat.CellValues[ColTypNiv,ARow] = 'SAL') then
    begin
      // Rechercher si l'établissement du salarié est Alsace ou non
      Q := OpenSQL('SELECT ETB_REGIMEALSACE FROM ETABCOMPL'
        + ' LEFT JOIN SALARIES ON PSA_ETABLISSEMENT = ETB_ETABLISSEMENT'
        + ' WHERE PSA_SALARIE = "' + (GEltNat.CellValues[ColValNiv,ARow]) + '"',True);
      if not Q.Eof then
      begin
//        if (Q.FindField('ETB_REGIMEALSACE').AsString = 'X') and (ACol = ColMontantEuro) then
//          GridGriseCell(GEltNat, Acol, Arow, Canvas);
//        if (Q.FindField('ETB_REGIMEALSACE').AsString = '-') and (ACol = ColMontant) then
//          GridGriseCell(GEltNat, Acol, Arow, Canvas);
      end;
      Ferme(Q);
    end;
  end;  }
end;

procedure TOF_UTOFSAISIESYNELTNAT.GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if ARow <> 0 then
  begin
    ZoneSuivanteouOk(GEltNat, ACol, ARow, Cancel);
    GEltNat.SetFocus;
  end
  else
    GEltNat.Enabled := False;
end;

procedure TOF_UTOFSAISIESYNELTNAT.ZoneSuivanteOuOk(Grille: THGrid; ACol, ARow: Integer; var Cancel: boolean);
var
  Sens, ii: integer;
  OldEna: boolean;
  CodLibre : String;
begin
  OldEna := Grille.SynEnabled;
  Grille.SynEnabled := False;
  Sens := -1;
  if Grille.Row > ARow then Sens := 1 else if ((Grille.Row = ARow) and (ACol < Grille.Col)) then Sens := 1;
  if (sens = -1) and ((ACol = ColStatus) or (Grille.Col = ColStatus)) and (ARow = 1) then
  begin
    Grille.SynEnabled := OldEna;
    Cancel := TRUE;
    exit;
  end;
  ACol := Grille.Col;
  ARow := Grille.Row;
  ii := 0;
  if Grille is TControl then
    CodLibre := (Grille as TControl).Name
  else
    CodLibre := 'GELTNAT';
  while not ZoneAccessible(CodLibre, ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 1000 then Break;
    if (ACol = ColStatus) and (ARow = 1) then
    begin
      ACol := 1;
      break;
    end;
    if Sens = 1 then
    begin
      if ((ACol = Grille.ColCount - 1) and (ARow = Grille.RowCount - 1)) then
      begin
        ACol := Grille.FixedCols;
        ARow := 1;
        Break;
      end;
      if ACol < Grille.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := Grille.FixedCols;
      end;
    end else
    begin
      if ((ACol = Grille.FixedCols) and (ARow = 1)) then Break;
      if ACol > Grille.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := Grille.ColCount - 1;
      end;
    end;
  end;
  Grille.SynEnabled := OldEna;
end;

function TOF_UTOFSAISIESYNELTNAT.ZoneAccessible(CodLibre: String; var ACol, ARow: Longint): Boolean;
begin
  if CodLibre = 'GCODLIBRE' then
  begin
    if (ACol <> ColCLDatVal) and (ACol <> ColCLLibelleVal) then
    begin
      Result := False;
      exit;
    end;
  end
  else
  begin
    if (ACol <> ColDatVal) and (ACol <> ColMontantEuro) then // and (ACol <> ColMontant)
    begin
      Result := False;
      exit;
    end;
  end;
  Result := True;
end;

procedure TOF_UTOFSAISIESYNELTNAT.GrilleCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var
  sDatVal : String;
begin
  if ARow <> 0 then
  begin
    if (ACol = ColDatVal) then
    begin
      sDatVal := GEltNat.CellValues[ACol,ARow];
      if not IsValidDate(sDatVal) then
      begin
        Cancel := True;
        PGIBox(TraduireMemoire('La date que vous avez saisie n''est pas valide'));
        exit;
      end;
    end;

    if (ACol = ColMontantEuro) then // or (ACol = ColMontant)
      GEltNat.CellValues[ColStatus,ARow] := '#ICO#44'; //    GEltNat.CellValues[ColStatus,ARow] := '#ICO#28';
  end
  else
    GEltNat.Enabled := False;
end;

procedure TOF_UTOFSAISIESYNELTNAT.ValidationSaisie(Sender : TObject);
var
  TobEltNationaux, TobCodeLibres : Tob;
  TEN : Tob;
  TobEltDynDos, TobCodeLibresFille : Tob;
  TED : Tob;
  i : integer;
  NbLigNonMemorisees : integer;
  ACol,ARow:integer;
  tmp : boolean;
  St : String;
  Q : TQuery;
  TobParamSalarie, TFind : Tob;
begin
  // Faire perdre le focus de dessus la grille pour valider la ligne en cours si saisie en cours
  if not AucunEltNational then
  begin
    ACol := GEltNat.Col;
    ARow := GEltNat.Row;
    tmp := False;
    GrilleCellExit(GEltNat,ACol,ARow,tmp);

    TobEltNationaux := TOB.Create('ELTNATIONAUX',nil,-1);
    TEN := TOB.Create('ELTNATIONAUX', TobEltNationaux, -1);
    TobEltDynDos := TOB.Create('ELTNATIONDOS',nil,-1);
    TED := TOB.Create('ELTNATIONDOS', TobEltDynDos, -1);
    NbLigNonMemorisees := 0;
    for i := 1 to (GEltNat.RowCount - 1) do
    begin
      if (GEltNat.CellValues[ColStatus,i] = '#ICO#44') then
      begin
        if GEltNat.CellValues[ColPredefini,i] <> '' then
        begin
          TEN.PutValue('PEL_PREDEFINI', GEltNat.CellValues[ColPredefini,i]);
          TEN.PutValue('PEL_NODOSSIER', V_PGI.NoDossier);
          TEN.PutValue('PEL_CONVENTION', GEltNat.CellValues[ColConvention,i]);
          TEN.PutValue('PEL_CODEELT', GEltNat.CellValues[ColCodeElt,i]);
          TEN.PutValue('PEL_DATEVALIDITE', GEltNat.CellValues[ColDatVal,i]);
          TEN.PutValue('PEL_LIBELLE', GEltNat.CellValues[ColLibelle,i]);
          TEN.PutValue('PEL_THEMEELT', GEltNat.CellValues[ColTheme,i]);
          TEN.PutValue('PEL_MONTANT', 0);
          TEN.PutValue('PEL_MONTANTEURO',StrToFloat(GEltNat.CellValues[ColMontantEuro,i])); //PT6
          TEN.PutValue('PEL_MONETAIRE', GEltNat.CellValues[ColMonetaire,i]);
          TEN.PutValue('PEL_ABREGE', GEltNat.CellValues[ColAbrege,i]);
          TEN.PutValue('PEL_BLOCNOTE', GEltNat.CellValues[ColBlocNote,i]);
          TEN.PutValue('PEL_DECALMOIS', GEltNat.CellValues[ColDecalMois,i]);
          TEN.PutValue('PEL_REGIMEALSACE', GEltNat.CellValues[ColRegimeAlsace,i]);
          TEN.PutValue('PEL_DATEMODIF', Date);
          TEN.PutValue('PEL_DATECREATION', Date);
          TEN.PutValue('PEL_UTILISATEUR', V_PGI.User);
          TEN.InsertOrUpdateDB;
        end
        else
        begin
          TED.PutValue('PED_TYPENIVEAU', GEltNat.CellValues[ColTypNiv,i]);
          TED.PutValue('PED_VALEURNIVEAU', GEltNat.CellValues[ColValNiv,i]);
          TED.PutValue('PED_CODEELT', GEltNat.CellValues[ColCodeElt,i]);
          TED.PutValue('PED_DATEVALIDITE', GEltNat.CellValues[ColDatVal,i]);
          TED.PutValue('PED_LIBELLE', GEltNat.CellValues[ColLibelle,i]);
          TED.PutValue('PED_THEMEELT', GEltNat.CellValues[ColTheme,i]);
          TED.PutValue('PED_MONTANT', 0);
          TED.PutValue('PED_MONTANTEURO', StrToFloat(GEltNat.CellValues[ColMontantEuro,i])); //PT6
          TED.PutValue('PED_MONETAIRE', GEltNat.CellValues[ColMonetaire,i]);
          TED.PutValue('PED_ABREGE', GEltNat.CellValues[ColAbrege,i]);
          TED.PutValue('PED_BLOCNOTE', GEltNat.CellValues[ColBlocNote,i]);
          TED.PutValue('PED_DECALMOIS', GEltNat.CellValues[ColDecalMois,i]);
          TED.PutValue('PED_REGIMEALSACE', GEltNat.CellValues[ColRegimeAlsace,i]);
          TED.PutValue('PED_ETABLISSEMENT', '');
          TED.PutValue('PED_SALARIE', '');
          TED.PutValue('PED_CODEPOP', '');
          if (GEltNat.CellValues[ColTypNiv,i] = 'ETB') then
            TED.PutValue('PED_ETABLISSEMENT', GEltNat.CellValues[ColETABLISSEMENT,i]);
          if (GEltNat.CellValues[ColTypNiv,i] = 'SAL') then
            TED.PutValue('PED_SALARIE', GEltNat.CellValues[ColSALARIE,i]);
          if (GEltNat.CellValues[ColTypNiv,i] = 'POP') then
            TED.PutValue('PED_CODEPOP', GEltNat.CellValues[ColCODEPOP,i]);
          TED.PutValue('PED_DATEMODIF', Date);
          TED.PutValue('PED_DATECREATION', Date);
          TED.PutValue('PED_UTILISATEUR', V_PGI.User);
          TED.InsertOrUpdateDB;
        end;
        ExecuteSQL('DELETE FROM PGSYNELTNAT WHERE PEY_TYPUTI = "1"'
          + ' AND PEY_CODEELT = "' + GEltNat.CellValues[ColCodeElt,i] + '"'
          + ' AND PEY_DATVAL = "' + USDateTime(StrToDate(GEltNat.CellValues[ColDatVal,i])) + '"'
          + ' AND PEY_PREDEFINI = "' + GEltNat.CellValues[ColPredefini,i] + '"'
          + ' AND PEY_NODOSSIER = "' + GEltNat.CellValues[ColNoDossier,i] + '"'
          + ' AND PEY_CONVENTION = "' + GEltNat.CellValues[ColConvention,i] + '"'
          + ' AND PEY_TYPNIV = "' + GEltNat.CellValues[ColTypNiv,i] + '"'
          + ' AND PEY_VALNIV = "' + GEltNat.CellValues[ColValNiv,i] + '"' );
      end
      else
        NbLigNonMemorisees := NbLigNonMemorisees + 1;
    end;
  end;
  TobEltNationaux.Free;

  if not AucunEltDynamique then
  begin
    // Faire perdre le focus de dessus la grille pour valider la ligne en cours si saisie en cours
    ACol := GCodLibre.Col;
    ARow := GCodLibre.Row;
    tmp := False;
    GCodLibreCellExit(GCodLibre,ACol,ARow,tmp);

    TobParamSalarie := Tob.Create('PARAMSALARIE',Nil,-1) ;
    Q := opensql('SELECT PPP_CODTABL,PPP_PGINFOSMODIF,PPP_PGTYPEDONNE FROM PARAMSALARIE ' +
            ' WHERE ##PPP_PREDEFINI## AND PPP_PGTYPEINFOLS = "ZLS"',True) ;
    if not Q.eof then TobParamSalarie.LoadDetailDb('PARAMSALARIE','','',Q,False) ;
    Ferme(Q) ;

    TobCodeLibres := TOB.Create('PGHISTODETAIL',nil,-1);
    TobCodeLibresFille := TOB.Create('PGHISTODETAIL', TobCodeLibres, -1);
    for i := 1 to (GCodLibre.RowCount - 1) do
    begin
      if (GCodLibre.CellValues[ColCLStatus,i] = '#ICO#44') then
      begin
        if GCodLibre.CellValues[ColCLCodLibre,i] = '' then           //PT5
          TFind := TobParamSalarie.FindFirst(['PPP_PGINFOSMODIF'],[GCodLibre.CellValues[ColCLAbrege,i]],false) //PT5
        else
          TFind := TobParamSalarie.FindFirst(['PPP_CODTABL'],[GCodLibre.CellValues[ColCLCodLibre,i]],false);
        TobCodeLibresFille.PutValue('PHD_SALARIE', GCodLibre.CellValues[ColCLSALARIE,i]);
        TobCodeLibresFille.PutValue('PHD_ETABLISSEMENT', '');
        TobCodeLibresFille.PutValue('PHD_GUIDHISTO', AglGetGuid());
        TobCodeLibresFille.PutValue('PHD_ORDRE', 0);
        if assigned(TFind) then
        begin
          TobCodeLibresFille.PutValue('PHD_PGINFOSMODIF', TFind.GetValue('PPP_PGINFOSMODIF'));
          TobCodeLibresFille.PutValue('PHD_TYPEVALEUR', TFind.GetValue('PPP_PGTYPEDONNE'));
        end;
        TobCodeLibresFille.PutValue('PHD_PGTYPEHISTO', '003');
        TobCodeLibresFille.PutValue('PHD_ANCVALEUR', '');
        TobCodeLibresFille.PutValue('PHD_NEWVALEUR', GCodLibre.CellValues[ColCLLibelleVal,i]);
        if GCodLibre.CellValues[ColCLCodLibre,i] = '' then     //PT5
          TobCodeLibresFille.PutValue('PHD_TABLETTE', '')      //PT5
        else
          TobCodeLibresFille.PutValue('PHD_TABLETTE', 'PGCOMBOZONELIBRE');
        TobCodeLibresFille.PutValue('PHD_PGTYPEINFOLS', 'ZLS');
        TobCodeLibresFille.PutValue('PHD_DATEAPPLIC', GCodLibre.CellValues[ColCLDatVal,i]);
        TobCodeLibresFille.PutValue('PHD_TRAITEMENTOK', 'X');
        TobCodeLibresFille.PutValue('PHD_DATEFINVALID', iDate1900);
        TobCodeLibresFille.PutValue('PHD_DATEDEBVALID', iDate1900);
        TobCodeLibresFille.PutValue('PHD_TYPEBUDG', '');
        TobCodeLibresFille.PutValue('PHD_NUMAUG', 0);
        TobCodeLibresFille.PutValue('PHD_ANNEE', '');
        TobCodeLibresFille.PutValue('PHD_RUBRIQUE', '');
        TobCodeLibresFille.PutValue('PHD_INFOHISTO', '');
        TobCodeLibresFille.PutValue('PHD_COMMENTARIE', '');
        TobCodeLibresFille.PutValue('PHD_THEMESALARIE', '');
        if GCodLibre.CellValues[ColCLCodLibre,i] = '' then   //PT5
          TobCodeLibresFille.PutValue('PHD_CODTABL', '')     //PT5
        else
          TobCodeLibresFille.PutValue('PHD_CODTABL', GCodLibre.CellValues[ColCLCodLibre,i]);
        TobCodeLibresFille.PutValue('PHD_CODEPOP', '');
        TobCodeLibresFille.PutValue('PHD_POPULATION', '');
        TobCodeLibresFille.PutValue('PHD_DATEMODIF', Date);
        TobCodeLibresFille.PutValue('PHD_DATECREATION', Date);
        TobCodeLibresFille.PutValue('PHD_CREATEUR', V_PGI.User);
        TobCodeLibresFille.PutValue('PHD_UTILISATEUR', V_PGI.User);
        TobCodeLibresFille.InsertOrUpdateDB;
        if GCodLibre.CellValues[ColCLCodLibre,i] = '' then   //PT5
          St := 'DELETE FROM PGSYNELTNAT WHERE PEY_TYPUTI = "3"' +
            ' AND PEY_ABREGE = "' + GCodLibre.CellValues[ColCLAbrege,i] + '"' +
            ' AND PEY_DATVAL = "' + USDateTime(StrToDate(GCodLibre.CellValues[ColCLDatVal,i])) + '"' +
            ' AND PEY_PREDEFINI = "' + GCodLibre.CellValues[ColCLPredefini,i] + '"' +
            ' AND PEY_NODOSSIER = "' + GCodLibre.CellValues[ColCLNoDossier,i] + '"' +
            ' AND PEY_CONVENTION = ""' +
            ' AND PEY_TYPNIV = "' + GCodLibre.CellValues[ColCLTypNiv,i] + '"' +
            ' AND PEY_VALNIV = "' + GCodLibre.CellValues[ColCLValNiv,i] + '"'
        else
          St := 'DELETE FROM PGSYNELTNAT WHERE PEY_TYPUTI = "3"' +
            ' AND PEY_CODEELT = "' + GCodLibre.CellValues[ColCLCodLibre,i] + '"' +
            ' AND PEY_DATVAL = "' + USDateTime(StrToDate(GCodLibre.CellValues[ColCLDatVal,i])) + '"' +
            ' AND PEY_PREDEFINI = "' + GCodLibre.CellValues[ColCLPredefini,i] + '"' +
            ' AND PEY_NODOSSIER = "' + GCodLibre.CellValues[ColCLNoDossier,i] + '"' +
            ' AND PEY_CONVENTION = ""' +
            ' AND PEY_TYPNIV = "' + GCodLibre.CellValues[ColCLTypNiv,i] + '"' +
            ' AND PEY_VALNIV = "' + GCodLibre.CellValues[ColCLValNiv,i] + '"';
        ExecuteSQL(St);
      end
      else
        NbLigNonMemorisees := NbLigNonMemorisees + 1;
    end;
  end;
  TobCodeLibres.Free;
  FreeAndNil(TobParamSalarie);

  if NbLigNonMemorisees <> 0 then
  begin
    PGIBox(TraduireMemoire('Certaines lignes n''ont pas été validées car n''ayant pas été saisies'));
    GEltNat.ClearSelected;
    GCodLibre.ClearSelected;
    ChargerDonnees;
  end
  else
    Ecran.Close;
end;


procedure TOF_UTOFSAISIESYNELTNAT.GrilleColEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if GCodLibre.row > 0 then
  begin
    If (ou = 5) Then
    begin
      GCodLibre.ElipsisButton := True;
      TypeElipsis := 'ZLS';
    end
    Else
    begin
      GCodLibre.ElipsisButton := False;
      TypeElipsis := '';
    end;
  end
  else
    GCodLibre.Enabled := False;
end;

procedure TOF_UTOFSAISIESYNELTNAT.GrilleElipsisClick(Sender: TObject);
var
  St,StWhere:String;
  i:Integer;
  Q:TQuery;
//PT7  ConvSal,Etab : String; //PT2
//PT7  DateMax :TDateTime;
begin
  If TypeElipsis='ZLS' then
  begin
    i := GCodLibre.row;
    if GCodLibre.CellValues[3,i] <> '' then //PT5
    begin
      StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET '
               + GetPlusPGCOMBOZONELIBRE(StrToDateTime(GCodLibre.CellValues[7,i]), GCodLibre.CellValues[3,i], GCodLibre.CellValues[4,i]); //PT7

{//PT7

      //DEB PT2
      Q := OpenSQL('SELECT PSA_CONVENTION,PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE = "' + GCodLibre.CellValues[4,i] + '"',True,1);
      if not Q.Eof then
      begin
        ConvSal := Q.FindField('PSA_CONVENTION').AsString;
        Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;   //PT4
      end;
      Ferme(Q);
      //FIN PT2
      //PT1
      //DEB PT4
      //Vérifier s'il n'existe pas des valeurs pour l'établissement du salarié
      if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
      ' WHERE PTE_PREDEFINI="DOS" AND PTE_NODOSSIER="' + V_PGI.NoDossier + '"' +
      ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GCodLibre.CellValues[7,i]))+'"' +
      ' AND PTE_CODTABL="'+GCodLibre.CellValues[3,i]+'"' +
      ' AND PTE_NIVSAIS="ETB" AND PTE_VALNIV="' + Etab + '"') then
      begin
        St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+GCodLibre.CellValues[3,i]+'" '+
          ' AND PTE_PREDEFINI="DOS" AND PTE_NODOSSIER="' + V_PGI.NoDossier + '"' +
          ' AND PTE_NIVSAIS="ETB" AND PTE_VALNIV="' + Etab + '"' +
          ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GCodLibre.CellValues[7,i])) + '"';
        Q := OpenSQL(St,True,1);
        DateMax := iDate1900;
        if not Q.Eof then
          DateMax := Q.FindField('DTVALID').AsDateTime;
        Ferme(Q);
        StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET ' +
          ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+GCodLibre.CellValues[3,i]+'"' +
          ' AND PTD_PREDEFINI="DOS" AND PTD_NODOSSIER="' + V_PGI.NoDossier + '"' +
          ' AND PTD_NIVSAIS="ETB" AND PTD_VALNIV="' + Etab + '"';
      end
      //Vérifier s'il n'existe pas des valeurs pour le dossier en général
      else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
      ' WHERE PTE_PREDEFINI="DOS" AND PTE_NODOSSIER="' + V_PGI.NoDossier + '"' +
      ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GCodLibre.CellValues[7,i]))+'"' +
      ' AND PTE_CODTABL="'+GCodLibre.CellValues[3,i]+'"' +
      ' AND PTE_NIVSAIS="GEN"') then
      begin
        St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+GCodLibre.CellValues[3,i]+'" '+
          ' AND PTE_PREDEFINI="DOS" AND PTE_NODOSSIER="' + V_PGI.NoDossier + '"' +
          ' AND PTE_NIVSAIS="GEN"' +
          ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GCodLibre.CellValues[7,i]))+'"';
        Q := OpenSQL(St,True,1);
        DateMax := iDate1900;
        if not Q.Eof then
          DateMax := Q.FindField('DTVALID').AsDateTime;
        Ferme(Q);
        StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET ' +
          ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+GCodLibre.CellValues[3,i]+'"' +
          ' AND PTD_PREDEFINI="DOS" AND PTD_NODOSSIER="' + V_PGI.NoDossier + '"' +
          ' AND PTD_NIVSAIS="GEN"';
      end
      //Vérifier s'il n'existe pas des valeurs pour STD + convention
      else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
      ' WHERE PTE_PREDEFINI="STD"' +
      ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GCodLibre.CellValues[7,i]))+'"' +
      ' AND PTE_CODTABL="'+GCodLibre.CellValues[3,i]+'"' +
      ' AND PTE_NIVSAIS="CON" AND PTE_VALNIV="' + ConvSal + '"') then
      begin
        St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+GCodLibre.CellValues[3,i]+'" '+
          ' AND PTE_PREDEFINI="STD"' +
          ' AND PTE_NIVSAIS="CON" AND PTE_VALNIV="' + ConvSal + '"' +
          ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GCodLibre.CellValues[7,i]))+'"';
        Q := OpenSQL(St,True,1);
        DateMax := iDate1900;
        if not Q.Eof then
          DateMax := Q.FindField('DTVALID').AsDateTime;
        Ferme(Q);
        StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET ' +
          ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+GCodLibre.CellValues[3,i]+'"' +
          ' AND PTD_PREDEFINI="STD"' +
          ' AND PTD_NIVSAIS="CON" AND PTD_VALNIV="' + ConvSal + '"';
      end
      //Vérifier s'il n'existe pas des valeurs pour STD + Convention 000
      else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
      ' WHERE PTE_PREDEFINI="STD"' +
      ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GCodLibre.CellValues[7,i]))+'"' +
      ' AND PTE_CODTABL="'+GCodLibre.CellValues[3,i]+'"' +
      ' AND PTE_NIVSAIS="CON" AND PTE_VALNIV="000"') then
      begin
        St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+GCodLibre.CellValues[3,i]+'" '+
          ' AND PTE_PREDEFINI="STD"' +
          ' AND PTE_NIVSAIS="CON" AND PTE_VALNIV="000"' +
          ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GCodLibre.CellValues[7,i]))+'"';
        Q := OpenSQL(St,True,1);
        DateMax := iDate1900;
        if not Q.Eof then
          DateMax := Q.FindField('DTVALID').AsDateTime;
        Ferme(Q);
        StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET ' +
          ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+GCodLibre.CellValues[3,i]+'"' +
          ' AND PTD_PREDEFINI="STD"' +
          ' AND PTD_NIVSAIS="CON" AND PTD_VALNIV="000"';
      end
      //Vérifier s'il n'existe pas des valeurs pour STD + GEN
      else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' +
      ' WHERE PTE_PREDEFINI="STD"' +
      ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GCodLibre.CellValues[7,i]))+'"' +
      ' AND PTE_CODTABL="'+GCodLibre.CellValues[3,i]+'"' +
      ' AND PTE_NIVSAIS="GEN"') then
      begin
        St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+GCodLibre.CellValues[3,i]+'" '+
          ' AND PTE_PREDEFINI="STD" AND PTE_NIVSAIS="GEN"' +
          ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GCodLibre.CellValues[7,i]))+'"';
        Q := OpenSQL(St,True,1);
        DateMax := iDate1900;
        if not Q.Eof then
          DateMax := Q.FindField('DTVALID').AsDateTime;
        Ferme(Q);
        StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET ' +
          ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+GCodLibre.CellValues[3,i]+'"' +
          ' AND PTD_PREDEFINI="STD" AND PTD_NIVSAIS="GEN"';
      end
      //Vérifier s'il n'existe pas des valeurs pour CEG
      else
      begin
        St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT WHERE PTE_CODTABL="'+GCodLibre.CellValues[3,i]+'" '+
          ' AND PTE_PREDEFINI="CEG" AND PTE_NIVSAIS="GEN"' +
          ' AND PTE_DTVALID<="'+USDATETIME(StrToDateTime(GCodLibre.CellValues[7,i]))+'"';
        Q := OpenSQL(St,True,1);
        DateMax := iDate1900;
        if not Q.Eof then
          DateMax := Q.FindField('DTVALID').AsDateTime;
        Ferme(Q);
        StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET ' +
          ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+GCodLibre.CellValues[3,i]+'"' +
          ' AND PTD_PREDEFINI="CEG" AND PTD_NIVSAIS="GEN"';
      end;
  //    StWhere := ' SELECT PTD_VALCRIT1,PTD_LIBELLECODE,PTD_RESULTAT FROM TABLEDIMDET' +
  //      ' WHERE PTD_CODTABL="' + GCodLibre.CellValues[3,i] + '" AND PTD_DTVALID = ' +
  //      ' (SELECT MAX(PTD_DTVALID) AS PTD_DTVALID FROM TABLEDIMDET WHERE PTD_DTVALID <= "' + USDATETIME(StrToDateTime(GCodLibre.CellValues[7,i])) + '" AND PTD_CODTABL="' + GCodLibre.CellValues[3,i] + '")' +
  //      ' AND (PTD_NIVSAIS = "GEN" OR (PTD_NIVSAIS ="CON" AND PTD_VALNIV="' +ConvSal+ '")' +  //PT2
  //      ' OR (PTD_NIVSAIS ="CON" AND PTD_VALNIV="000"))'; //+  //PT2
  //PT3      ' ORDER BY PTD_LIBELLECODE';
      //FIN PT4

}
      LookUpList (GCodLibre,'Eléments dynamiques','','PTD_VALCRIT1','','','PTD_LIBELLECODE',TRUE,-1,StWhere); //PT3
      If GCodLibre.CellValues[5,i]<>'' Then
      begin
        St := 'PTD_CODTABL = "' + GCodLibre.CellValues[3,i] + '"';
        Q := OpenSQL('SELECT PTD_LIBELLECODE FROM TABLEDIMDET WHERE '+St+' AND PTD_VALCRIT1="'+GCodLibre.CellValues[5,i]+'"',True);
        If not Q.eof then
          GCodLibre.CellValues[6,i] := Q.FindField('PTD_LIBELLECODE').AsString;
        Ferme(Q);
      end;
    end;
  end;
end;

procedure TOF_UTOFSAISIESYNELTNAT.GCodLibreCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if ARow <> 0 then
  begin
    ZoneSuivanteOuOk(GCodLibre, ACol, ARow, Cancel);
    GCodLibre.SetFocus;
  end
  else
    GCodLibre.Enabled := False;
end;

procedure TOF_UTOFSAISIESYNELTNAT.GCodLibreCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  sDatVal : String;
begin
  if ARow <> 0 then
  begin
    if (ACol = ColCLDatVal) then
    begin
      sDatVal := GCodLibre.CellValues[ACol,ARow];
      if not IsValidDate(sDatVal) then
      begin
        Cancel := True;
        PGIBox(TraduireMemoire('La date que vous avez saisie n''est pas valide'));
        exit;
      end;
    end;

    if (ACol = ColCLLibelleVal) then
      GCodLibre.CellValues[ColCLStatus,ARow] := '#ICO#44';
  end
  else
    GCodLibre.Enabled := False;
end;

procedure TOF_UTOFSAISIESYNELTNAT.GCodLibrePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ARow <> 0 then
  begin
    If   (ACol = ColCLLibelle) or (ACol = ColCLCodLibre) or (ACol = ColCLSalarie) or (ACol = ColCLValeur)
      or (ACol = ColCLStatus) or AucunEltDynamique then
        GridGriseCell(GCodLibre, Acol, Arow, Canvas);
  end
  else
    GCodLibre.Enabled := False;
end;

Initialization
  registerclasses ( [ TOF_UTOFSAISIESYNELTNAT ] ) ;
end.
