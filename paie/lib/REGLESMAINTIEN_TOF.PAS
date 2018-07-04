{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/07/2004
Modifié le ... : 27/07/2004
Description .. : Source TOF de la FICHE : REGLESMAINTIEN ()
Suite ........ : paramétrage des différents seuils (durée) , nbre de jours de
Suite ........ : carence, taux à appliquer
Mots clefs ... : TOF;REGLESMAINTIEN
*****************************************************************}
{
PT1     30/11/2004 MF V_60 Correction traitement maintien qd champ catégorie renseigné.
PT2     20/12/2004 PH V_60 Erreur SQL DB2
}
Unit REGLESMAINTIEN_TOF ;

Interface

Uses
//unused    StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
//unused     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
//unused     mul,
{$else}
//unused     eMul,
{$ENDIF}
     EntPaie,
//unused     forms,
     sysutils,
//unused     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     P5Util,
     PgOutils,
     UTOB,
     UTOF ;

Type
  TOF_REGLESMAINTIEN = Class (TOF)
  private
    CEG, STD, DOS                                                     : Boolean;
    LCodeMaint, LLibelle, LConvention, LCateg, LPredefini, LNoDossier : string;
    LAnciennete , MaxOrdre ,  LigneError                              : integer;
    GLesRegles                                                        : THGrid;
    TOB_Regles                                                        : TOB;
    BtnVal                                                            : TToolBarButton97;
    procedure ChargementDesRegles();
    procedure GLesReglesCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GLesReglesCellExit(Sender: TObject; var  ACol, ARow: Integer; var Cancel: Boolean);
    procedure GLesReglesRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GLesReglesRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BTnDelClick(Sender: TObject);
    procedure BTnValClick(Sender: TObject);

  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_REGLESMAINTIEN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_REGLESMAINTIEN.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_REGLESMAINTIEN.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 27/07/2004
Modifié le ... : 27/07/2004
Description .. : procédure OnLoad
Suite ........ : Chargement des données de la fiche, principalement celles 
Suite ........ : de la grille
Mots clefs ... : PAIE; MAINTIEN
*****************************************************************}
procedure TOF_REGLESMAINTIEN.OnLoad ;
begin
  Inherited ;
  ChargementDesRegles
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 27/07/2004
Modifié le ... : 27/07/2004
Description .. : procédure OnArgument
Suite ........ : chargement de la fiche. 
Suite ........ : récupération des paramètres passée par 
Suite ........ : CRITMAINTIEN_TOF
Suite ........ : Paramétrage de boutons
Suite ........ : Paramétrage de la grille de saisie
Mots clefs ... : PAIE; MAINTIEN
*****************************************************************}
procedure TOF_REGLESMAINTIEN.OnArgument (S : String ) ;
var
  st, Categ                               : string;
  BtnDel                                  : TToolBarButton97;

begin
  Inherited ;
  AccesPredefini('TOUS', CEG, STD, DOS);

  st := Trim(S);
  LCodeMaint := ReadTokenSt(st);
  LLibelle := ReadTokenSt(st);
  LConvention := ReadTokenSt(st);
  LCateg := ReadTokenSt(st);
  LAnciennete := StrToInt(ReadTokenSt(st));
  LPredefini := ReadTokenSt(st);
  LNoDossier := ReadTokenSt(st);
  Categ := ReadTokenSt(st);

  // paramétrage bouton suppression
  BtnDel := TToolBarButton97(GetControl('BDelete'));
  if BtnDel <> nil then
    BtnDel.OnClick := BTnDelClick;
  // paramétrage bouton valider
  BtnVal := TToolBarButton97(GetControl('BValider'));
  if BtnVal <> nil then
    BtnVal.OnClick := BTnValClick;

// d PT1
  SetControlText('LCATEG', '');
  SetControlProperty('LIBCATEG','Caption',Categ+' '+LCateg);
// f PT1
  if (VH_Paie.PGCritMaintien = '') then
  begin
    setControlVisible('LCATEG',False);
    setControlVisible('LIBCATEG',False);
  end;

  SetControlText('LCODEMAINT',LCodeMaint);
  SetControlText('LLIBELLE',LLibelle);
  SetControlText('LCONVENTION',RechDom ('PGCONVENTIONS',LConvention,FALSE));
{PT1  if (VH_Paie.PGCritMaintien <> '') or  (Categ <> '')then
    SetControlText('LCATEG', LCateg);
    SetControlText('LCATEG',RechDom (Categ,LCateg,FALSE));}
  SetControlText('LANCIENNETE',IntToStr(LAnciennete)+' mois');

  GLesRegles := THGrid(Getcontrol('LESREGLES'));

  if GLesRegles <> nil then
  // Grille des règles de maintien
  begin
    // mise en place des propriétés d'affichage des cellules
    GLesRegles.ColAligns[0] := taRightJustify;
    GLesRegles.ColEditables[0] := False;
    GLesRegles.ColAligns[1] := taRightJustify;
    GLesRegles.ColAligns[2] := taRightJustify;
    GLesRegles.ColFormats[3] := '# ##0.00';
    GLesRegles.ColAligns[3] := taRightJustify;
    GLesRegles.OnCellEnter := GLesReglesCellEnter;
    GLesRegles.OnCellExit := GLesReglesCellExit;
//    GLesRegles.OnDblClick := GLesReglesDblClick;
    GLesRegles.OnKeyDown :=  KeyDown;
    GLesRegles.OnRowEnter := GLesReglesRowEnter;
    GLesRegles.OnRowExit := GLesReglesRowExit;

   end
  else exit;

  if (LPredefini = 'CEG') then
  begin
    SetControlEnabled('BDelete', CEG);
    GLesRegles.ColEditables[1] := CEG;
    GLesRegles.ColEditables[2] := CEG;
    GLesRegles.ColEditables[3] := CEG;
    BtnVal.enabled := CEG;
  end;

  if (LPredefini = 'STD') then
  begin
    SetControlEnabled('BDelete', STD);
    GLesRegles.ColEditables[1] := STD;
    GLesRegles.ColEditables[2] := STD;
    GLesRegles.ColEditables[3] := STD;
    BtnVal.enabled := STD;
end;

  if (LPredefini = 'DOS') then
  begin
    SetControlEnabled('BDelete', DOS);
    GLesRegles.ColEditables[1] := DOS;
    GLesRegles.ColEditables[2] := DOS;
    GLesRegles.ColEditables[3] := DOS;
    BtnVal.enabled := DOS;
  end;

end ;

procedure TOF_REGLESMAINTIEN.OnClose ;
begin
  Inherited ;
  FreeAndNil(TOB_Regles);
end ;

procedure TOF_REGLESMAINTIEN.OnDisplay () ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 27/07/2004
Modifié le ... :   /  /    
Description .. : Procédure OnCancel
Suite ........ : Annulation de la saisie : raz des lignes + Rechargement des 
Suite ........ : données initiales
Mots clefs ... : PAIE; MAINTIEN
*****************************************************************}
procedure TOF_REGLESMAINTIEN.OnCancel () ;
var
  i                             : integer;
begin
  Inherited ;
  for i := 2 to GLesRegles.RowCount do
  begin
      GLesRegles.Cells[0, i] := '';
      GLesRegles.Cells[1, i] := '';
      GLesRegles.Cells[2, i] := '';
      GLesRegles.Cells[3, i] := '';
  end;
  GLesRegles.FixedRows := 1;
  // on recharge le contenu de la fiche
  OnLoad;
end ;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 27/07/2004
Modifié le ... : 27/07/2004
Description .. : procédure ChargementDesRegles
Suite ........ : chargement des lignes de la grille de saisie à partir des 
Suite ........ : éléments de REGLESMAINTIEN correspondant à la table 
Suite ........ : Critères sélectionnée
Mots clefs ... : PAIE ; MAINTIEN
*****************************************************************}
procedure TOF_REGLESMAINTIEN.ChargementDesRegles();
var
  QL                                    : TQuery;
  LesRegles                             : TOB;
  ligne                                 : integer;
begin
  // Création de la TOB_Reglts
  TOB_Regles := TOB.Create('Les règles', nil, -1);

  QL := OpenSql('SELECT *'+
    ' FROM REGLESMAINTIEN ' +
    'WHERE ' +
    'PAM_CODEMAINT="'+LCodeMaint+'" AND '+
    'PAM_PREDEFINI="'+LPredefini+'" AND '+
    'PAM_NODOSSIER="'+LNoDossier+'" AND '+
    'PAM_CONVENTION="'+LConvention+'" AND '+
    'PAM_BORNEFINANC='+IntToStr(LAnciennete), TRUE); // PT2

  TOB_Regles.LoadDetailDB('REGLESMAINTIEN', '', '', QL, FALSE, FALSE);
  Ferme(QL);
  GLesRegles.RowCount := TOB_Regles.Detail.Count + 1 ;
  MaxOrdre := 0;

  if TOB_Regles <> nil then
  // chargement de la grille des règles
  begin
    LesRegles := TOB_Regles.FindFirst([''], [''], TRUE);
    Ligne := 1;
    while LesRegles <> nil do
    begin
      if (LesRegles.GetValue('PAM_ORDRE') > MaxOrdre) then
        MaxOrdre := StrToInt(LesRegles.GetValue('PAM_ORDRE'));
      GLesRegles.Cells[0, Ligne] := LesRegles.GetValue('PAM_ORDRE');
      GLesRegles.Cells[1, Ligne] := StrfMontant(LesRegles.GetValue('PAM_CARENCE'),15,0,'', TRUE);
      GLesRegles.Cells[2, Ligne] := StrfMontant(LesRegles.GetValue('PAM_NBJMAINTIEN'),15,0,'', TRUE);
      GLesRegles.Cells[3, Ligne] := DoubleToCell(LesRegles.GetValue('PAM_TXMAINTIEN'),2);
      if (LesRegles.GetValue('PAM_TXMAINTIEN') = 0) then
        GLesRegles.Cells[3, Ligne] := '0,00';

      Ligne := Ligne + 1;
      LesRegles := TOB_Regles.FindNext([''], [''], TRUE);
    end;
    if (TOB_Regles.Detail.Count = 0) then
    // aucun règlement --> affichage d'une ligne vide
    begin
      GLesRegles.RowCount := 2;
      GLesRegles.FixedRows := 1;
      GLesRegles.Cells[0, 1] := IntToStr(MaxOrdre+1);
      GLesRegles.Cells[1, 1] := '';
      GLesRegles.Cells[2, 1] := '';
      GLesRegles.Cells[3, 1] := '';
      MaxOrdre := 1;
    end;
  end;
end; {fin ChargementDesRegles}
procedure TOF_REGLESMAINTIEN.GLesReglesCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
end; {fin GLesReglesCellEnter}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 27/07/2004
Modifié le ... : 27/07/2004
Description .. : procédure GLesReglesCellExit
Suite ........ : sortie de cellule
Suite ........ : vérif ligne OK
Mots clefs ... : PAIE; MAINTIEN
*****************************************************************}
procedure TOF_REGLESMAINTIEN.GLesReglesCellExit(Sender: TObject; var  ACol, ARow: Integer; var Cancel: Boolean);
var
   Taux                                                 : double;
begin
  if (Acol = 3) then
  begin
    Taux :=  Valeur(GLesRegles.Cells[3, ARow]);
    GLesRegles.Cells[3, ARow] :=  DoubleToCell(Taux,2);
    if (Taux = 0) then
      GLesRegles.Cells[3, ARow] := '0,00';
  end;

  if (GLeSRegles.RowCount-1 = Arow+1) and (GLesRegles.Cells[0,ARow+1] = '')  then
  begin
    MaxOrdre := MaxOrdre+1;
    GLesRegles.Cells[0,ARow+1] := IntToStr(MaxOrdre);
  end;

  if (GLesRegles.Cells[1,ARow] = '') or
     (GLesRegles.Cells[2,ARow] = '') or
     (GLesRegles.Cells[3,ARow] = '') then
  begin
    LigneError := Arow;
    BtnVal.enabled := false;
  end
  else
  if (LigneError = Arow) then
  begin
    LigneError := 0;
    BtnVal.enabled := true;
  end;

end; { fin GLesReglesCellExit}
procedure TOF_REGLESMAINTIEN.GLesReglesRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
end; {fin  GLesReglesRowEnter}
procedure TOF_REGLESMAINTIEN.GLesReglesRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
end; { fin GLesReglesRowExit}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 27/07/2004
Modifié le ... : 27/07/2004
Description .. : procédure KeyDown
Suite ........ : Quand flèche en bas , ouverture d'une nouvelle ligne si bas
Suite ........ : de grille
Mots clefs ... : PAIE; MAINTIEN
*****************************************************************}
procedure TOF_REGLESMAINTIEN.KeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ((key = 40) and (GLesRegles.row = GLesRegles.RowCount-1)) then
   if((LPredefini = 'CEG') and (CEG = True) or
      (LPredefini = 'STD') and (STD = True) or
      (LPredefini = 'DOS') and (DOS = True)) then 
    // flèche bas
    if (GLesRegles.Cells[0,GLesRegles.RowCount-1] <> '') and
       (GLesRegles.Cells[1,GLesRegles.RowCount-1] <> '') and
       (GLesRegles.Cells[2,GLesRegles.RowCount-1] <> '') and
       (GLesRegles.Cells[3,GLesRegles.RowCount-1] <> '') then
    begin
      LigneError := GLesRegles.RowCount;
      GLesRegles.RowCount := GLesRegles.RowCount+1;
      BtnVal.enabled := false;
    end;
end;  {fin KeyDown}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 27/07/2004
Modifié le ... : 27/07/2004
Description .. : procédure BTnDelClick
Suite ........ : suppression d'un ligne de la grille
Mots clefs ... : PAIE; MAINTIEN
*****************************************************************}
procedure TOF_REGLESMAINTIEN.BTnDelClick(Sender: TObject);
var
  reponse                               : integer;
begin
  if (GLesRegles.Cells[1,GLesRegles.row] = '') or
     (GLesRegles.Cells[2,GLesRegles.row] = '') or
     (GLesRegles.Cells[3,GLesRegles.row] = '') then
  // règlement non saisi
  begin
    if (StrToInt(GLesRegles.Cells[0,GLesRegles.row])  = MaxOrdre) then
      MaxOrdre := MaxOrdre-1;
    if (LigneError = GLesRegles.row) then
    begin
      LigneError := 0;
      BtnVal.enabled := true;
    end;
    LigneError := LigneError - 1;
    GLesRegles.DeleteRow(GLesRegles.row);
  end
  else
  begin
   reponse := PGIAsk(' Confirmez-vous la suppression de la ligne ?',
                          ' règle');
   if (reponse=mrYes) then
   begin
     if (StrToInt(GLesRegles.Cells[0,GLesRegles.row])  = MaxOrdre) then
       MaxOrdre := MaxOrdre-1;
     if (LigneError = GLesRegles.row) then
     begin
       LigneError := 0;
       BtnVal.enabled := true;
     end;
     LigneError := LigneError-1;
     GLesRegles.DeleteRow(GLesRegles.row);
   end;
  end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 27/07/2004
Modifié le ... : 27/07/2004
Description .. : procédure BTnValClick
Suite ........ : Validation de la grille. 
Suite ........ : les lignes d'origine sont d'abord supprimées puis insertion 
Suite ........ : des lignes de la grille
Mots clefs ... : PAIE; MAINTIEN
*****************************************************************}
procedure TOF_REGLESMAINTIEN.BTnValClick(Sender: TObject);
var
  st                                            : string;
  Ligne                                         : integer;
  LesRegles                                     : TOB;
begin
  try
    begintrans;
    if TOB_Regles <> nil then TOB_Regles.SetAllModifie(TRUE);
    // suppression des éléments de la table REGLESMAINTIEN affichés au chargement
    // de la fiche

    st := 'DELETE'+
    ' FROM REGLESMAINTIEN ' +
    'WHERE ' +
    'PAM_CODEMAINT="'+LCodeMaint+'" AND '+
    'PAM_PREDEFINI="'+LPredefini+'" AND '+
    'PAM_NODOSSIER="'+LNoDossier+'" AND '+
    'PAM_CONVENTION="'+LConvention+'" AND '+
    'PAM_BORNEFINANC='+IntToStr(LAnciennete);  // PT2

    ExecuteSQL(st);

    FreeAndNil(TOB_Regles);

    TOB_Regles := TOB.Create('Les règles', nil, -1);
    // Création de la TOB_Regles à partir des lignes de la grille
    for ligne := 1 to GLesRegles.RowCount-1 do
    begin
      LesRegles := TOB.create('REGLESMAINTIEN', TOB_Regles, -1);
      LesRegles.PutValue('PAM_CODEMAINT',LCodeMaint);
      LesRegles.PutValue('PAM_PREDEFINI',LPredefini);
      LesRegles.PutValue('PAM_NODOSSIER',LNoDossier);
      LesRegles.PutValue('PAM_CONVENTION',LConvention);
      LesRegles.PutValue('PAM_BORNEFINANC',LAnciennete);
      LesRegles.PutValue('PAM_ORDRE',ligne );
      LesRegles.PutValue('PAM_CARENCE',Valeur(GLesRegles.Cells[1,ligne]));
      LesRegles.PutValue('PAM_NBJMAINTIEN',Valeur(GLesRegles.Cells[2,Ligne]));
      LesRegles.PutValue('PAM_TXMAINTIEN',Valeur(GLesRegles.Cells[3,Ligne]));
    end;
    TOB_Regles.InsertDB(nil, FALSE);

    Committrans;

  except
    Rollback;
    PGIBox('! Erreur maj table REGLESMAINTIEN ', '');
  end;

end;

Initialization
  registerclasses ( [ TOF_REGLESMAINTIEN ] ) ;
end.
