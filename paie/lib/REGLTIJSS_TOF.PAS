{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/05/2004
Modifié le ... : 02/06/2004
Description .. : Source TOF de la FICHE : REGLTIJSS ()
Suite ........ : Saisie, modification et suppression des règlements d'IJSS
Suite ........ : avec pointage (dépointage) des absences .
Mots clefs ... : TOF;REGLTIJSS
*****************************************************************}
{
PT1  07/02/2005 MF V_6.0 : Ajout d'un champ commentaire sur la grille
                                des absences ( utilisation de PCN_LIBCOMPL2)
PT2  25/04/2005 MF V_602 : 1- FQ 12083  Modification des dates d'absence
                           d'1 règlement déja intégré
                           2- FQ 12169  Mouvement absence non soldé à tort.   
PT3  23/01/2006 SB V_65 FQ 10866 Ajout clause predefini motif d'absence
PT4  19/06/2006 SB V_65 FQ 13231 Retrait des mvt absences annulées
PT5  16/04/2007 MF V_702 FQ 14111 accès au suivi des règlements IJSS via la
                         fiche salarié (ribbon bar)
PT6  22/10/2007 MF V_80 FQ 14111 inti de la grille des réglements
PT7  21/12/2007 FC V_81 Concept accessibilité salarié
}
Unit REGLTIJSS_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
//unused     mul,
{$else}
//unused     eMul,
{$ENDIF}
//unused     forms,
     sysutils,
//unused     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     ParamDat,
//PGVisuObjet,     
     P5Util,
     UTOB,
     UTOF ;

Type
  TOF_REGLTIJSS = Class (TOF)
    public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    private
    WdateDeb, WDateFin, DateNaissance                   : TDateTime;
    Matricule,Etablissement,NomSalarie,PrenomSalarie    : string;
    NoSS                                                : string;
    GAbsences, GReglt                                   : THGrid;
    IJSSNonSoldees                                      : TCheckBox;
    TOB_Absences,TOB_Reglt                              : TOB;
    Chgt                                                : integer;
    BtnCherch                                           : TToolBarButton97; // PT5
    ActionParam:String;//PT7

    procedure DateElipsisclick(Sender: TObject);
    procedure GRegltDblClick(Sender: TObject);
    procedure IJSSNonSoldeesClick(Sender: TObject);
    procedure ChargementGAbsences(IJSSNonSoldees : boolean);
    procedure ChargementGReglt();
    procedure KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GRegltCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GRegltCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GRegltRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GRegltRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    function  LigneModifiee(Ou : Integer): Boolean;
    function  ChgLigneOrigine(Ou : Integer): Boolean;
    function  LigneOK(Ou : Integer): Boolean;
    procedure NouveauReglement(Ou: Integer);
    procedure ModifReglement(Ou: Integer);
    procedure BTnDelClick(Sender: TObject);
    procedure BTnValClick(Sender: TObject);
    procedure GAbsencesExit(Sender: TObject);
//    procedure GAbsencesEnter(Sender: TObject);
    procedure GAbsencesCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);

    procedure GAbsencesDblClick(Sender: TObject);
    procedure BTnChercheClick (Sender: TObject); // PT5

  end ;

Type TLReglt = record
    DateReglt           : string;
    MtRegle             : string;
    DateDebutAbs        : string;
    DateFinAbs          : string;
    NbjCarence          : string;
    NbjIjss             : string;
    MtJournalier        : string;
    MtIjssBrutes        : string;
    TypeAbs             : string;
end;
Type LigneAction = (Modification, Creation);

var
  LigneOrigine                                  : TLReglt;
  NbjIJSSPrevus                                 : double;
  NbjIJSSRegles                                 : double;
  TotIJSSNettes                                 : double;
  Action                                        : LigneAction;
  LigneValide, LigneVide,LigneExisteDeja        : boolean;
  LigneError                                    : integer;
  BtnVal                                        : TToolBarButton97;
Implementation

procedure TOF_REGLTIJSS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_REGLTIJSS.OnDelete ;
begin
  Inherited ;
end ;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... : 02/06/2004
Description .. : procédure OnLoad
Suite ........ : Chargements de la grille des absences et de la grille de
Suite ........ : saisie des règlements.
Suite ........ : La saisie débute sur la grille des règlements
Suite ........ : Affichage des différents cumuls
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGLTIJSS.OnLoad ;
begin
  Inherited ;

  SetControlText('PCN_SALARIE',Matricule);
  SetControlText('NOMPRENOM',Trim(NomSalarie)+'  '+Trim(PrenomSalarie));
  SetControlText('ABSDU',DateToStr(WDateDeb));
  SetControlText('ABSAU',DateToStr(WDateFin));
  SetControlText('IJSSSOLDEES','-');
  SetControlText('PSA_NUMEROSS',NoSS);
  if (DateNaissance <> IDate1900) then
    SetControlText('PSA_DATENAISSANCE',DateToStr(DateNaissance));

  // Chargement des grilles
  BtnCherch.Click;    // PT5

end ; {fin OnLoad}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... : 02/06/2004
Description .. : procédure OnArgument
Suite ........ : Chargement de la fiche
Suite ........ : Mise en palce des différents propriétés d'affichage
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.OnArgument (S : String ) ;
var
  BtnDel                        : TToolBarButton97;
  FicheSal                      : string; // PT5
begin
  Inherited ;

// d PT5
// accès à partir de la fiche salarié
  FicheSal := '';
  if (Pos('FICHESAL', S) > 0) then
  begin
    FicheSal := ReadTokenSt(S);
  end;
// f PT5

  Matricule := ReadTokenSt(S);
  Etablissement := ReadTokenSt(S);
  NomSalarie := ReadTokenSt(S);
  PrenomSalarie := ReadTokenSt(S);
  WdateDeb := StrToDate(ReadTokenSt(S));
  WDateFin := StrToDate(ReadTokenSt(S));
  NoSS := ReadTokenSt(S);
  DateNaissance := StrToDate(ReadTokenSt(S));
  ActionParam := '';
  if S<>'' then
    ActionParam := ReadTokenSt(S); //PT7

// d PT5
// accès à partir de la fiche salarié
  if (FicheSal = 'FICHESAL') then
  begin
    WDateFin := Date;
    WDateDeb := PlusMois(WDateFin,-12);
    SetControlProperty ('ABSDU','Enabled',True);
    SetControlProperty ('ABSAU','Enabled',True);
    SetControlProperty ('BCherche','Visible',True);
  end
  else
  begin
    SetControlProperty ('ABSDU','Enabled',false);
    SetControlProperty ('ABSAU','Enabled',false);
    SetControlProperty ('BCherche','Visible',false);
  end;
  BtnCherch:= TToolBarButton97(GetControl('BCherche'));
  if BtnCherch <> nil then
    BtnCherch.OnClick := BTnChercheClick;
// f PT5

  // Grille des absences
  GAbsences := THGrid(Getcontrol('GABSENCES'));
  // Grille des règlements
  GReglt := THGrid(Getcontrol('GREGLT'));

  // paramétrage bouton suppression
  BtnDel := TToolBarButton97(GetControl('BDelete'));
  if BtnDel <> nil then
    BtnDel.OnClick := BTnDelClick;

  BtnVal := TToolBarButton97(GetControl('BValider'));
  if BtnVal <> nil then
    BtnVal.OnClick := BTnValClick;

  //DEB PT7
  if ActionParam='ACTION=CONSULTATION' then
  begin
    SetControlProperty ('GABSENCES','Enabled',false);
    SetControlProperty ('GREGLT','Enabled',false);
    SetControlProperty ('BValider','Visible',false);
  end;
  //FIN PT7

  // checkBox "IJSS non soldées" invisible et inaccessible - procédure à revoir
  SetControlProperty('IJSSSOLDEES','Visible',FALSE);
  SetControlProperty('IJSSSOLDEES','Enabled',FALSE);
  IJSSNonSoldees := TCheckBox(Getcontrol('IJSSSOLDEES'));
  If IJSSNonSoldees <> NIL then
   IJSSNonSoldees.OnClick := IJSSNonSoldeesClick;

  if GAbsences <> nil then
  // Grille des absences
  begin
    // mise en place des propriétés d'affichage des cellules
    GAbsences.ColTypes[0] := 'D';
    GAbsences.ColFormats[0] := ShortDateFormat;
    GAbsences.ColAligns[0] := taCenter;
    GAbsences.ColEditables[0] := false;
    GAbsences.ColTypes[1] := 'D';
    GAbsences.ColFormats[1] := ShortDateFormat;
    GAbsences.ColAligns[1] := taCenter;
    GAbsences.ColEditables[1] := false;
    GAbsences.ColFormats[2] := 'CB=PGMOTIFABSENCE';
    GAbsences.ColAligns[2] := taCenter;
    GAbsences.ColEditables[2] := false;
    GAbsences.ColAligns[3] := taRightJustify;
    GAbsences.ColEditables[3] := false;
    GAbsences.ColAligns[4] := taRightJustify;
    GAbsences.ColEditables[4] := false;
    GAbsences.ColAligns[5] := taRightJustify;
    GAbsences.ColEditables[5] := false;
    GAbsences.ColAligns[6] := taRightJustify;
    GAbsences.ColEditables[6] := false;
    GAbsences.ColTypes[7] := 'B';
    GAbsences.ColFormats[7] := IntToStr(Ord(csCoche));
    GAbsences.ColEditables[7] := false;
    GAbsences.ColAligns[7] := taCenter;
// d PT1
    GAbsences.ColAligns[8] := taLeftJustify;
    GAbsences.ColEditables[8] := true;
// f PT1    

    GAbsences.OnExit := GAbsencesExit ;
//    GAbsences.OnEnter := GAbsencesEnter ;
    GAbsences.OnDblClick := GAbsencesDblClick;
    GAbsences.OnCellEnter := GAbsencesCellEnter ;
    GAbsences.col := 7;   // seule la cellule "soldée" est accessible
    GAbsences.FlipBool:=True;
  end
  else exit;

  if GReglt <> nil then
  // Grille des règlements d'IJSS
  begin
    // mise en place des propriétés d'affichage des cellules
    GReglt.ColTypes[0] := 'D';
    GReglt.ColFormats[0] := ShortDateFormat;
    GReglt.ColAligns[0] := taCenter;

    GReglt.ColTypes[1] := 'F';
    GReglt.ColAligns[1] := taRightJustify;
    GReglt.ColTypes[2] := 'D';
    GReglt.ColFormats[2] := ShortDateFormat;
    GReglt.ColAligns[2] := taCenter;
    GReglt.ColTypes[3] := 'D';
    GReglt.ColFormats[3] := ShortDateFormat;
    GReglt.ColAligns[3] := taCenter;
    GReglt.ColAligns[4] := taRightJustify;
    GReglt.ColAligns[5] := taRightJustify;
    GReglt.ColAligns[6] := taRightJustify;
    GReglt.ColAligns[7] := taRightJustify;
    GReglt.ColFormats[8] := 'CB=PGTYPEABS';
    GReglt.ColTypes[9] := 'D';
    GReglt.ColFormats[9] := ShortDateFormat;
    GReglt.ColAligns[9] := taCenter;
    GReglt.ColEditables[9] := false;
    GReglt.OnCellEnter := GRegltCellEnter;
    GReglt.OnCellExit := GRegltCellExit;
    GReglt.OnDblClick := GRegltDblClick;
    GReglt.OnKeyDown :=  KeyDown;
    GReglt.OnRowEnter := GRegltRowEnter;
    GReglt.OnRowExit := GRegltRowExit;

  end
  else exit;
end ; {fin OnArgument}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... :   /  /
Description .. : Procédure OnClose
Suite ........ : Sortie de la fiche
Suite ........ : Libération des TOB
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.OnClose ;
begin
  Inherited ;
  FreeAndNil(TOB_Absences);
  FreeAndNil(TOB_Reglt);
end ; {fin OnClose}

procedure TOF_REGLTIJSS.OnDisplay () ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... : 02/06/2004
Description .. : procédure OnCancel
Suite ........ : correspond à "défaire"
Suite ........ : on vide les grille et on relance OnLoad (Chargement du
Suite ........ : contenu de la fiche)
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.OnCancel () ;
var
  i                             : integer;
begin
  Inherited ;
  // on vide la grille absences
  for i := 2 to GAbsences.RowCount do
  begin
      GAbsences.Cells[0, i] := '';
      GAbsences.Cells[1, i] := '';
      GAbsences.Cells[2, i] := '';
      GAbsences.Cells[3, i] := '';
      GAbsences.Cells[4, i] := '';
      GAbsences.Cells[5, i] := '';
      GAbsences.Cells[6, i] := '';
      GAbsences.Cells[7, i] := '';
      GAbsences.Cells[8, i] := ''; //  PT1
  end;

  // on vide la grille règlement
  for i := 2 to GReglt.RowCount do
  begin
      GReglt.Cells[0, i] := '';
      GReglt.Cells[1, i] := '';
      GReglt.Cells[2, i] := '';
      GReglt.Cells[3, i] := '';
      GReglt.Cells[4, i] := '';
      GReglt.Cells[5, i] := '';
      GReglt.Cells[6, i] := '';
      GReglt.Cells[7, i] := '';
      GReglt.CellValues[8, i] := '';
      GReglt.Cells[9, i] := '';
  end;

  GAbsences.FixedRows := 1;
  GReglt.FixedRows := 1;

  // on recharge le contenu de la fiche
  OnLoad;
end ; {fin OnCancel}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... :   /  /
Description .. : procédure DateElipsisclick
Suite ........ : Choix date sur calendrier
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.DateElipsisclick(Sender: TObject);
var
  key : char;
begin
  key := '*';
  ParamDate (Ecran, Sender, Key);
end;  {fin DateElipsisclick}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... :   /  /
Description .. : procédure GRegltDBLClick
Suite ........ : positionne l'elipsis uniquement sur les cellules "dates" de la
Suite ........ : grille règlements
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.GRegltDblClick(Sender: TObject);
begin
  if ((GReglt.col = 0) or (GReglt.col = 2) or (GReglt.col = 3)) and
     (GReglt.row <> 0) then
    DateElipsisclick(Sender);
end; {fin GRegltDblClick}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... :   /  /
Description .. : procédure GRegltCellEnter
Suite ........ : positionne l'elipsis uniquement sur les cellules "dates" de la
Suite ........ : grille règlements
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.GRegltCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if ((GReglt.col = 0) or (GReglt.col = 2) or (GReglt.col = 3)) and
     (GReglt.row <> 0) then
    GReglt.ElipsisButton := TRUE
  else
    GReglt.ElipsisButton := FALSE;
// d PT2-1   FQ  12083
   if(GReglt.Cells[9,GReglt.row] <> '') and
     ((GReglt.col = 2) or (GReglt.col = 3))then
           PGIBox('Règlement déjà intégré dans un bulletin. '+
                  'La modification de cette date est déconseillée', '');
// f PT2-1
end; {fin GRegltCellEnter}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... : 02/06/2004
Description .. : procédure GAbsencesExit
Suite ........ : En sortie de la grille des absences on met à jour la
Suite ........ : TOB_Absences, le champs PCN_IJSSSOLDEE
Suite ........ : uniquement.
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.GAbsencesExit(Sender: TObject);
var
  LesAbsences                           : TOB;
  ligne                                 : integer;
begin
  for ligne := 1 to GAbsences.RowCount-1 do
  begin
    LesAbsences := TOB_Absences.FindFirst(['LIGNE'],[ligne],TRUE);
    if (LesAbsences <> nil) then
    begin
      LesAbsences.PutValue('PCN_IJSSSOLDEE', GAbsences.Cells[7, ligne]);
      LesAbsences.PutValue('PCN_LIBCOMPL2', GAbsences.Cells[8, ligne]); // PT1 PT2-2 
    end;
  end;
end;  {fin GAbsencesExit}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... :   /  /
Description .. : procédure GAbsencesCellEnter
Suite ........ : sur la grille des absences on entre systématiquement sur la
Suite ........ : cellule "soldée"
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.GAbsencesCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if (GAbsences.col = 7) then
    GAbsences.FlipBool:=True
  else
    GAbsences.FlipBool:=False;

end; {fin GAbsencesCellEnter}

{***********A.G.L.***********************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... : 02/06/2004
Description .. : procédure IJSSNonSoldeesClick
Suite ........ :
Suite ........ : !! inutilisée pour l'instant car recharge la grille des absences
Suite ........ : à partir de la table sans tenir compte des modifications en
Suite ........ : cours
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.IJSSNonSoldeesClick(Sender: TObject);
begin
if IJSSNonSoldees <> NIL then
   ChargementGAbsences (IJSSNonSoldees.Checked);
end;  {fin IJSSNonSoldeesClick}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... : 02/06/2004
Description .. : procédure ChargementGAbsences
Suite ........ : création de la TOB_Absences et chargement de la grille des
Suite ........ : absences
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.ChargementGAbsences(IJSSNonSoldees : boolean);
var
  QL                                    : TQuery;
  st                                    : string;
  LesAbsences                           : TOB;
  ligne                                 : integer;
begin
  NbjIJSSPrevus := 0;

  // on charge initialement toutes les absences soldées ou non
  if (IJSSNonSoldees = TRUE) then
    st := 'PCN_IJSSSOLDEE = "-" AND '
  else
    st := '';

  // création de la TOB_Absences
  TOB_Absences := TOB.Create('Les lignes d''absence', nil, -1);

  QL := OpenSql('SELECT PCN_SALARIE,PCN_DATEDEBUTABS,PCN_DATEFINABS,'+
    'PCN_TYPECONGE,PCN_JOURS,PCN_JOURS,PCN_NBJCALEND,PCN_NBJCARENCE,'+
    'PCN_NBJIJSS,PCN_IJSSSOLDEE,PCN_ORDRE,PCN_LIBCOMPL2'+             // PT1
    ' FROM ABSENCESALARIE ' +
    'WHERE ' + st+
    '(PCN_GESTIONIJSS = "X") AND '+
    '(PCN_SALARIE ="' + Matricule + '") AND ' +
    '(PCN_DATEDEBUTABS >="' + UsDateTime(WDateDeb) + '") AND ' +
    '(PCN_DATEDEBUTABS <="' + UsDateTime(WDateFin) + '") AND ' +
    '(PCN_DATEFINABS >= "' + UsDateTime(WDateDeb) + '") AND ' +
    '(PCN_DATEFINABS <= "' + UsDateTime(WDateFin) + '") '+
    'AND PCN_ETATPOSTPAIE <> "NAN" '+ { PT4 }
    'ORDER BY PCN_DATEFINABS', TRUE);

  TOB_Absences.LoadDetailDB('ABSENCESALARIE', '', '', QL, FALSE, FALSE);
  Ferme(QL);
  GAbsences.RowCount := TOB_Absences.Detail.Count + 1;
  if TOB_Absences <> nil then
  // chargement de la grille des absences
  begin
    LesAbsences := TOB_Absences.FindFirst([''], [''], TRUE);

    if (LesAbsences = nil) then
    // cas d'une TOB vide (affichage d'au moins une ligne à blanc)
    begin
      GAbsences.Cells[0, 1] := '';
      GAbsences.Cells[1, 1] := '';
      GAbsences.Cells[2, 1] := '';
      GAbsences.Cells[3, 1] := '';
      GAbsences.Cells[4, 1] := '';
      GAbsences.Cells[5, 1] := '';
      GAbsences.Cells[6, 1] := '';
      GAbsences.Cells[7, 1] := '';
      GAbsences.Cells[8, 1] := '';              // PT1
      GAbsences.RowCount :=  GAbsences.RowCount + 1;
      GAbsences.FixedRows := 1;
    end
    else
    begin
      Ligne := 1;
      while LesAbsences <> nil do
      begin
        LesAbsences.AddChampSupValeur('LIGNE',Ligne, FALSE);
        GAbsences.Cells[0, Ligne] := DateToStr(LesAbsences.GetValue('PCN_DATEDEBUTABS'));
        GAbsences.Cells[1, Ligne] := DateToStr(LesAbsences.GetValue('PCN_DATEFINABS'));
        GAbsences.CellValues[2, Ligne] := LesAbsences.GetValue('PCN_TYPECONGE');
        GAbsences.Cells[3, Ligne] := StrfMontant(LesAbsences.GetValue('PCN_JOURS'),15,0,'',TRUE);
        GAbsences.Cells[4, Ligne] := StrfMontant(LesAbsences.GetValue('PCN_NBJCALEND'),15,0,'',TRUE);
        GAbsences.Cells[5, Ligne] := StrfMontant(LesAbsences.GetValue('PCN_NBJCARENCE'),15,0,'',TRUE);
        GAbsences.Cells[6, Ligne] := StrfMontant(LesAbsences.GetValue('PCN_NBJIJSS'),15,0,'',TRUE);
        GAbsences.Cells[7, Ligne] := LesAbsences.GetValue('PCN_IJSSSOLDEE');
        GAbsences.Cells[8, Ligne] := LesAbsences.GetValue('PCN_LIBCOMPL2');  // PT1

        Ligne := Ligne + 1;

        // Cumul du nombre de jours d'IJSS prévus
        NbjIJSSPrevus := NbjIJSSPrevus + LesAbsences.GetValue('PCN_NBJIJSS');
        LesAbsences := TOB_Absences.FindNext([''], [''], TRUE);
      end;
    end;
  end;
end; { fin ChargementGAbsences }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... : 02/06/2004
Description .. : procédure ChargementGReglt
Suite ........ : création de la TOB_Reglts et chargement de la grille des
Suite ........ : régelements
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.ChargementGReglt();
var
  QL                                    : TQuery;
  LesReglts                             : TOB;
  ligne                                 : integer;
begin
    NbjIJSSRegles := 0;
    TotIJSSNettes := 0.00;
//d PT6
    for ligne := 1 to  Greglt.RowCount do
    begin
      Greglt.Cells[0, ligne] := '';
      Greglt.Cells[1, ligne] := '';
      Greglt.Cells[2, ligne] := '';
      Greglt.Cells[3, ligne] := '';
      Greglt.Cells[4, ligne] := '';
      Greglt.Cells[5, ligne] := '';
      Greglt.Cells[6, ligne] := '';
      Greglt.Cells[7, ligne] := '';
      Greglt.CellValues[8, ligne] := '';
      Greglt.Cells[9, ligne] := '';
    end;
//f PT6
  // Création de la TOB_Reglts
  TOB_Reglt := TOB.Create('Les lignes de règlements', nil, -1);

  QL := OpenSql('SELECT *'+
    ' FROM REGLTIJSS ' +
    'WHERE ' +
    '(PRI_SALARIE ="' + Matricule + '") AND ' +
    '(PRI_DATEDEBUTABS >="' + UsDateTime(WDateDeb) + '") AND ' +
    '(PRI_DATEDEBUTABS <="' + UsDateTime(WDateFin) + '") AND ' +
    '(PRI_DATEFINABS >= "' + UsDateTime(WDateDeb) + '") AND ' +
    '(PRI_DATEFINABS <= "' + UsDateTime(WDateFin) + '") '+
    'ORDER BY PRI_DATEFINABS', TRUE);

  TOB_Reglt.LoadDetailDB('REGLTIJSS', '', '', QL, FALSE, FALSE);
  Ferme(QL);
  Greglt.RowCount := TOB_Reglt.Detail.Count + 2;
{if V_PGI.Debug=True then
   PGVisuUnObjet( TOB_Reglt,'','') ;}

  if TOB_Reglt <> nil then
  // chargement de la grille des règlements
  begin
    LesReglts := TOB_Reglt.FindFirst([''], [''], TRUE);
    Ligne := 1;
    while LesReglts <> nil do
    begin
      Greglt.Cells[0, Ligne] := DateToStr(LesReglts.GetValue('PRI_DATEREGLT'));
      Greglt.Cells[1, Ligne] := StrfMontant(LesReglts.GetValue('PRI_MTREGLE'),15,2,'', TRUE);
      Greglt.Cells[2, Ligne] := DateToStr(LesReglts.GetValue('PRI_DATEDEBUTABS'));
      Greglt.Cells[3, Ligne] := DateToStr(LesReglts.GetValue('PRI_DATEFINABS'));
      Greglt.Cells[4, Ligne] := StrfMontant(LesReglts.GetValue('PRI_NBJCARENCE'), 15, 0, '', TRUE);
      Greglt.Cells[5, Ligne] := StrfMontant(LesReglts.GetValue('PRI_NBJIJSS'),15,0,'', TRUE);
      Greglt.Cells[6, Ligne] := StrfMontant(LesReglts.GetValue('PRI_MTJOURNALIER'),15,2,'', TRUE);
      Greglt.Cells[7, Ligne] := StrfMontant(LesReglts.GetValue('PRI_MTIJSSBRUTES'),15,2,'', TRUE);
      Greglt.CellValues[8, Ligne] := LesReglts.GetValue('PRI_TYPEABS');
      if (LesReglts.GetValue('PRI_DATEINTEGR') = IDate1900) then
        Greglt.Cells[9,Ligne ] := ''
      else
        Greglt.Cells[9, Ligne] := DateToStr(LesReglts.GetValue('PRI_DATEINTEGR'));
      Ligne := Ligne + 1;
      NbjIJSSRegles := NbjIJSSRegles + LesReglts.GetValue('PRI_NBJIJSS');
      TotIJSSNettes := TotIJSSNettes + LesReglts.GetValue('PRI_MTREGLE');
      LesReglts := TOB_Reglt.FindNext([''], [''], TRUE);
    end;
    if (TOB_Reglt.Detail.Count = 0) then
    // aucun règlement --> affichage d'une ligne vide
    begin
      GReglt.FixedRows := 1;
      Greglt.Cells[0, 1] := '';
      Greglt.Cells[1, 1] := '';
      Greglt.Cells[2, 1] := '';
      Greglt.Cells[3, 1] := '';
      Greglt.Cells[4, 1] := '';
      Greglt.Cells[5, 1] := '';
      Greglt.Cells[6, 1] := '';
      Greglt.Cells[7, 1] := '';
      Greglt.CellValues[8, 1] := '';
      Greglt.Cells[9, 1] := '';
    end;
  end;
end; {fin ChargementGReglt}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... :   /  /
Description .. : procédure KeyDown
Suite ........ : quand on fait flèche en bas sur la dernière ligne de la grille
Suite ........ : des règlements on ouvre une ligne vide
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ((key = 40) and (GReglt.row = GReglt.RowCount-1) and
      (LigneValide = TRUE)) then
    // flèche bas
    if (GReglt.Cells[0,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[1,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[2,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[3,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[4,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[5,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[6,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[7,GReglt.RowCount-1] <> '') and
       (GReglt.CellValues[8,GReglt.RowCount-1] <> '') then
    begin
      GReglt.RowCount := GReglt.RowCount+1;
    end;
end;  {fin KeyDown}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... :   /  /    
Description .. : procédure GRegltCellExit
Suite ........ : sortie de cellule sur la grille des règlements
Suite ........ : contrôle validuité dates
Suite ........ : formatage des doubles
Suite ........ : qd on sort de la dernière cellule de la dernière ligne on 
Suite ........ : ouvre une nouvelle ligne
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.GRegltCellExit(Sender: TObject; var  ACol, ARow: Integer; var Cancel: Boolean);
var
  NbDec                         : integer;
  Valzone                       : double;
begin
  // contrôle validité des dates
  if ((Acol = 0) or (Acol = 2) or (Acol = 3)) and
      (Arow <> 0) then
    if (GReglt.Cells[Acol,Arow] <> '')and
       (GReglt.Cells[Acol,Arow] <> '  /  /    ') and
       not (IsValidDate(GReglt.Cells[Acol,Arow])) then
    begin
      PGIBox('Date erronée', '');

      GReglt.col := Acol;
      GReglt.row := Arow;
    end;

 // contrôle validité des dates
  if ((Acol = 3)) and
      (Arow <> 0) then
    if (GReglt.Cells[Acol,Arow] <> '')and
       (GReglt.Cells[Acol,Arow] <> '  /  /    ')and
       (StrToDate(GReglt.Cells[Acol,Arow]) < StrToDate(GReglt.Cells[2,Arow])) then
    begin
      PGIBox('Date de fin erronée', '');

      GReglt.col := Acol;
      GReglt.row := Arow;
    end;

  // formatage des doubles
  if ((Acol <> 0) and (Acol <> 2) and (Acol <> 3)) and
      (Arow <> 0) and (Acol <> 8) and (Acol <> 9) then
  begin
    ValZone := Valeur(GReglt.Cells[Acol, Arow]);
    if ((Acol = 4) or (Acol =5)) then
    begin
      NbDec := 0;
      Valzone := Int(Valzone);
    end
    else
      NbDec := 2;
    if (ValZone <> 0) then
      GReglt.Cells[Acol,Arow] := DoubleToCell(ValZone, NbDec)
    else
      if (NbDec = 0) then
        GReglt.Cells[Acol,Arow] :=  '0'
      else
        GReglt.Cells[Acol,Arow] := '0,00';
  end;

  // si sortie de la dernière cellule de la denière ligne
  if ((Acol = 8) and (GReglt.row = GReglt.RowCount-1) and
      (LigneValide = TRUE)) then
    // ouverture d'une ligne à blanc
    if (GReglt.Cells[0,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[1,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[2,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[3,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[4,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[5,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[6,GReglt.RowCount-1] <> '') and
       (GReglt.Cells[7,GReglt.RowCount-1] <> '') and
       (GReglt.CellValues[8,GReglt.RowCount-1] <> '') then
    begin
      GReglt.RowCount := GReglt.RowCount+1;
      GReglt.Row :=  GReglt.RowCount-1;
      GReglt.Col := 0;
      Chgt := 0;    // indicateur utiliser pour ne pas passer 2 fois RowExit
    end;
end; { fin GRegltCellExit}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... : 02/06/2004
Description .. : procédure GRegltRowExit
Suite ........ : sortie d'une ligne de la grille des règlements.
Suite ........ : traitements contrôle de la ligne, de création et de 
Suite ........ : modification
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGLTIJSS.GRegltRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
 if (Chgt = 0) then
    // on vient de créer une ligne à blanc, il ne faut pas passer 2 fois dans
    // le RoxExit
    Chgt := 1
  else
  begin
    LigneValide := True;
    if (LigneModifiee(Ou) = TRUE) or (Action = Creation) then
    // la ligne a été modifiée ou il s'agit d'un nouveau règlement
    begin
      if  not(LigneOK(Ou)) and (LigneVide = False) then
      // la ligne n'est pas valide
      begin
        LigneValide := False;
        if (LigneModifiee(Ou) = TRUE) and (GReglt.RowCount <> Ou) then
        begin
          if (Ou <> GReglt.RowCount-1) then
            GReglt.row := Ou+1
          else
            GReglt.row := Ou;
        end;
      end
      else
      // la ligne est valide
      begin
        if (LigneVide = False) then
        begin
          LigneValide := True;
          if (Action = Creation) then
          // nouvelle ligne
          begin
            NouveauReglement(Ou);
            if (LigneValide = False) then
            // règlement non pris en compte
            begin
              if (Ou <> GReglt.RowCount-1) then
                GReglt.row := Ou+1
              else
                GReglt.row := Ou;
            end
            else
            // règlement OK  - Cumuls nbre de jours réglés et montant IJSS nettes
            begin
              if (GReglt.Cells[5,Ou] <> '') and(GReglt.Cells[5,Ou] <> '0') then
                NbjIJSSRegles := NbjIJSSRegles + Valeur(GReglt.Cells[5,Ou]);
              if (GReglt.Cells[1,Ou] <> '') and(GReglt.Cells[5,Ou] <> '0,00') then
                TotIJSSNettes := TotIJSSNettes + Valeur(GReglt.Cells[1,Ou]);
            end;
          end
          else
          // Modification de ligne
          begin
            ModifReglement(Ou);
            if (LigneValide = False) then
            //  règlement non pris en compte
            begin
{              if (Ou <> GReglt.RowCount-1) then
                GReglt.row := Ou+1
              else}
                GReglt.row := Ou;
            end
            else
            // règlement OK  - Cumuls nbre de jours réglés et montant IJSS nettes
            begin
              NbjIJSSRegles := NbjIJSSRegles - Valeur(LigneOrigine.NbjIjss);
              TotIJSSNettes := TotIJSSNettes - Valeur(LigneOrigine.MtRegle);
              if (GReglt.Cells[5,Ou] <> '') and (GReglt.Cells[5,Ou] <> '0') then
                NbjIJSSRegles := NbjIJSSRegles + Valeur(GReglt.Cells[5,Ou]);
              if (GReglt.Cells[1,Ou] <> '') and (GReglt.Cells[1,Ou] <> '0,00') then
                TotIJSSNettes := TotIJSSNettes + Valeur(GReglt.Cells[1,Ou]);
            end;
          end; {Modification de ligne}
        end; {if (LigneVide = False)}
      end; {la ligne est valide}

      // affichage des cumuls
      SetControlProperty('TOTIJSS', 'Value', TotIJSSNettes);
      SetControlProperty('JOURSPREV', 'Value', NbjIJSSPrevus);
      SetControlProperty('JOURSREGL', 'Value', NbjIJSSRegles);
      SetControlProperty('ECART', 'Value',(NbjIJSSPrevus-NbjIJSSRegles));

    end; {la ligne a été modifiée ou il s'agit d'un nouveau règlement}

    if ((LigneModifiee(Ou) = False) and
       (LigneValide = False) and
       (Action = Modification))  then
    begin
      if (Ou <> GReglt.RowCount-1) then
        GReglt.row := Ou+1
      else
        GReglt.row := Ou;
    end;

    if (LigneModifiee(Ou) = False) and
       (LigneValide = True) and
       (LigneError <> 0) then
    // correction date d'un règlements mentionné comme déjà existant
    // on a remis les dates d'origine
    begin
      LigneExisteDeja := False;
      LigneError := 0;
      BtnVal.Enabled := True;
    end;
  end;
end; {fin GRegltRowExit}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... :   /  /    
Description .. : procédure GRegltRowEnter
Suite ........ : Stockage ligne avant modification
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGLTIJSS.GRegltRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  ChgLigneOrigine(Ou);
  if (LigneOK(Ou)) and (LigneVide = False)then
    LigneValide := True
  else
    LigneValide := False;
end; {fin GRegltRowEnter}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... : 02/06/2004
Description .. : procédure LigneModifiee
Suite ........ : contrôle s'il s'agit d'une création de ligne  ou d'une 
Suite ........ : modification et si le contenu de la ligne a effectivement été 
Suite ........ : modifié
Mots clefs ... : 
*****************************************************************}
function TOF_REGLTIJSS.LigneModifiee(Ou : Integer): Boolean;
begin
  if ((LigneOrigine.DateReglt = '') or
      (LigneOrigine.DateReglt = '  /  /    ')) and
     ((LigneOrigine.MtRegle = '') or
      (LigneOrigine.MtRegle = '0,00')) and
     ((LigneOrigine.DateDebutAbs = '') or
      (LigneOrigine.DateDebutAbs = '  /  /    ')) and
     ((LigneOrigine.DateFinAbs = '')  or
      (LigneOrigine.DateFinAbs = '  /  /    ')) and
     ((LigneOrigine.NbjCarence = '') or
      (LigneOrigine.NbjCarence = '0'))and
     ((LigneOrigine.NbJIjss = '') or
      (LigneOrigine.NbJIjss = '0')) and
     ((LigneOrigine.MtJournalier = '') or
      (LigneOrigine.MtJournalier = '0,00'))and
     ((LigneOrigine.MtIjssBrutes = '') or
      (LigneOrigine.MtIjssBrutes = '0,00'))  then
     Action :=  Creation
    else
     Action :=  Modification;

     if (LigneOrigine.DateReglt <> Greglt.Cells[0, Ou]) or
       (LigneOrigine.MtRegle <> Greglt.Cells[1, Ou]) or
       (LigneOrigine.DateDebutAbs <> Greglt.Cells[2, Ou]) or
       (LigneOrigine.DateFinAbs <>  Greglt.Cells[3, Ou]) or
       (LigneOrigine.NbjCarence <>  Greglt.Cells[4, Ou]) or
       (LigneOrigine.NbJIjss <>   Greglt.Cells[5, Ou]) or
       (LigneOrigine.MtJournalier <> Greglt.Cells[6, Ou]) or
       (LigneOrigine.MtIjssBrutes <> Greglt.Cells[7, Ou]) or
       (LigneOrigine.TypeAbs <> GReglt.CellValues[8,Ou]) then
      result := TRUE
    else
      result := FALSE;
end;  {fin LigneModifiee}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... :   /  /    
Description .. : procédure ChgLigneOrigine
Suite ........ : sauvegarde contenu de la ligne sur laquelle on va travailler
Mots clefs ... : 
*****************************************************************}
function TOF_REGLTIJSS.ChgLigneOrigine(Ou : Integer): Boolean;
begin
    result := TRUE;
    if (Ou = 0) then  exit;
    LigneOrigine.DateReglt := Greglt.Cells[0, Ou];
    LigneOrigine.MtRegle := Greglt.Cells[1, Ou];
    LigneOrigine.DateDebutAbs := Greglt.Cells[2, Ou];
    LigneOrigine.DateFinAbs :=  Greglt.Cells[3, Ou];
    LigneOrigine.NbjCarence :=  Greglt.Cells[4, Ou];
    LigneOrigine.NbJIjss :=   Greglt.Cells[5, Ou];
    LigneOrigine.MtJournalier := Greglt.Cells[6, Ou];
    LigneOrigine.MtIjssBrutes := Greglt.Cells[7, Ou];
    LigneOrigine.TypeAbs:= Greglt.CellValues[8, Ou];

    result := TRUE;
end; {fin ChgLigneOrigine}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... :   /  /
Description .. : fonction LigneOK
Suite ........ : vérification de la validité de la ligne de règlement
Suite ........ : obligatoire : date de règlement, dates d'absence, montant 
Suite ........ : IJSS nettes, nbre de jours IJSS réglés
Mots clefs ... : 
*****************************************************************}
function TOF_REGLTIJSS.LigneOK(Ou : Integer): Boolean;
begin
  LigneVide := False;
  if (Greglt.Cells[0, Ou] = '') or (Greglt.Cells[0, Ou] = '  /  /    ') or
     (Greglt.Cells[1, Ou] = '') or (Greglt.Cells[1, Ou] = '0,00')  or
     (Greglt.Cells[2, Ou] = '') or (Greglt.Cells[2, Ou] = '  /  /    ') or
     (Greglt.Cells[3, Ou] = '') or (Greglt.Cells[3, Ou] = '  /  /    ') or
     (Greglt.Cells[5, Ou] = '') or (Greglt.Cells[5, Ou] = '0') then
    result := FALSE
  else
    result := TRUE;

  if ((Greglt.Cells[0, Ou] = '') or (Greglt.Cells[0, Ou] = '  /  /    ')) and
     ((Greglt.Cells[1, Ou] = '') or (Greglt.Cells[1, Ou] = '0,00'))  and
     ((Greglt.Cells[2, Ou] = '') or (Greglt.Cells[2, Ou] = '  /  /    ')) and
     ((Greglt.Cells[3, Ou] = '') or (Greglt.Cells[3, Ou] = '  /  /    ')) and
     ((Greglt.Cells[5, Ou] = '') or (Greglt.Cells[5, Ou] = '0'))then
  // remise à blanc de la ligne car aucune date et aucun montant saisis
  begin
      GReglt.Cells[0, Ou] := '';
      GReglt.Cells[1, Ou] := '';
      GReglt.Cells[2, Ou] := '';
      GReglt.Cells[3, Ou] := '';
      GReglt.Cells[4, Ou] := '';
      GReglt.Cells[5, Ou] := '';
      GReglt.Cells[6, Ou] := '';
      GReglt.Cells[7, Ou] := '';
      GReglt.CellValues[8, Ou] := '';
      GReglt.Cells[9, Ou] := '';
      LigneVide := TRUE;
  end;
end;  {fin LigneOK}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... : 02/06/2004
Description .. : procédure NouveauReglement
Suite ........ : vérif si règlement pas déjà saisi
Suite ........ : traitement de pointage des absences
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGLTIJSS.NouveauReglement(Ou: Integer);
var
  LesAbsences                           : TOB;
  reponse, i                            : integer;
  begin
  LigneExisteDeja := False;
  // Recherche si règlement déjà existant pour ces dates d'absence
  // la recherche se fait sur les lignes de la grille, donc tient
  // compte des lignes Ajoutées, Supprimées, Modifiées.
  for i:= 1 to GReglt.RowCount-1 do
  begin
    if (i <> Ou) then
      if (Greglt.Cells[2,i] <> '') and (Greglt.Cells[2,i] <> '  /  /    ') and
         (Greglt.Cells[3,i] <> '') and (Greglt.Cells[3,i] <> '  /  /    ') and
         (((StrToDate(GReglt.Cells[2,i]) <= StrToDate(GReglt.Cells[2,Ou])) and
          (StrToDate(GReglt.Cells[3,i]) >= StrToDate(GReglt.Cells[2,Ou]))) or
         ((StrToDate(GReglt.Cells[2,i]) >= StrToDate(GReglt.Cells[2,Ou])) and
          (StrToDate(GReglt.Cells[2,i]) <= StrToDate(GReglt.Cells[3,Ou])))) then
      // un règlement pour ces dates d'absence existe déjà
      begin
        PGIBox ('Un règlement pour ces dates d''absence existe déjà','Saisie règlement');
        LigneValide := False;
        LigneExisteDeja := True;
        LigneError := Ou;
        BtnVal.Enabled := False;
        
        if (Ou = GReglt.RowCount-2) then
          GReglt.RowCount := GReglt.RowCount-1;
        GReglt.Row :=  GReglt.RowCount-1;
        GReglt.Col := 0;
        break;
      end;
  end;
  if (LigneValide = False) then exit;

  if ((GReglt.Cells[2,Ou] <> '') and (GReglt.Cells[2,Ou] <> '  /  /    ')) or
     ((GReglt.Cells[3,Ou] <> '') and (GReglt.Cells[3,Ou] <> '  /  /    ')) then
  begin
  // Recherche s'il existe un mouvement pour ces dates d'absence
  LesAbsences := TOB_Absences.FindFirst(['PCN_DATEDEBUTABS','PCN_DATEFINABS'],
                                        [StrToDate(GReglt.Cells[2,Ou]),
                                         StrToDate(GReglt.Cells[3,Ou])],
                                        TRUE);
  if (LesAbsences <> nil) then
  // Une absence pour cette date existe
  begin
    if (GReglt.Cells[5,Ou] = LesAbsences.GetValue('PCN_NBJIJSS')) then
     // le nbre de jours d'IJSS réglés = nbre de jours IJSS prévus
     // On solde la ligne d'absence
    begin
      LesAbsences.PutValue('PCN_IJSSSOLDEE', 'X');
      GAbsences.Cells[7, LesAbsences.GetValue('LIGNE')] := LesAbsences.GetValue('PCN_IJSSSOLDEE');
      PGIInfo ('Un mouvement d''absence soldé','Pointage');
    end
    else
    begin
      reponse := PGIAsk('Un mouvement d''absence pour ces dates existe mais le#10#13'+
                       ' nombre de jours IJSS prévu diffère du nombre de jours#10#13'+
                       ' IJSS réglé.#10#13'+
                       ' Voulez-vous solder le mouvement ?','Pointage');
      if (reponse=mrYes) then
        // On solde la ligne d'absence
        LesAbsences.PutValue('PCN_IJSSSOLDEE', 'X');
        GAbsences.Cells[7, LesAbsences.GetValue('LIGNE')] := LesAbsences.GetValue('PCN_IJSSSOLDEE');
    end;
  end
  else
    PGIInfo('Aucun mouvement d''absence ne correspond à ce règlement','Pointage');
  end;
end; {fin NouveauReglement}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... :   /  /    
Description .. : procédure ModifReglement
Suite ........ : vérif. si après modif, ce règlement ne fait pas doublons.
Suite ........ : traitement du pointage des absences
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGLTIJSS.ModifReglement(Ou: Integer);
var
  reponse, i                            : integer;
  LesAbsences                           : TOB;

begin
  if (LigneError <> 0) then
  // correction date d'un règlement mentionné comme déjà existant
  begin
    LigneExisteDeja := False;
    LigneError := 0;
    BtnVal.Enabled := True;
  end;

  if (LigneOrigine.DateDebutAbs <> Greglt.Cells[2, Ou]) or
     (LigneOrigine.DateFinAbs <>  Greglt.Cells[3, Ou]) then
  // Modif des dates d'absence
  begin
    // Recherche si règlement déjà existant pour ces dates d'absence
    // la recherche se fait sur les lignes de la grille, donc tient
    // compte des lignes Ajoutées, Supprimées, Modifiées.
    for i:= 1 to GReglt.RowCount-1 do
    begin
      if (i <> Ou) then
        if (Greglt.Cells[2,i] <> '') and (Greglt.Cells[2,i] <> '  /  /    ') and
           (Greglt.Cells[3,i] <> '') and (Greglt.Cells[3,i] <> '  /  /    ') and
           (((StrToDate(GReglt.Cells[2,i]) <= StrToDate(GReglt.Cells[2,Ou])) and
             (StrToDate(GReglt.Cells[3,i]) >= StrToDate(GReglt.Cells[2,Ou]))) or
            ((StrToDate(GReglt.Cells[2,i]) >= StrToDate(GReglt.Cells[2,Ou])) and
             (StrToDate(GReglt.Cells[2,i]) <= StrToDate(GReglt.Cells[3,Ou])))) then
        // un règlement pour ces dates d'absence existe déjà
        begin
          PGIBox ('Un règlement pour ces dates d''absence existe déjà','Saisie règlement');
          LigneValide := False;
          LigneExisteDeja := True;
          LigneError := Ou;
          BtnVal.Enabled := False;
          GReglt.Col := 0;
          break;
        end;
    end;
    if (LigneValide = False) then exit;

    // Recherche s'il existe un mouvement pour ces dates d'absence (avant modification)
    LesAbsences := TOB_Absences.FindFirst(['PCN_DATEDEBUTABS','PCN_DATEFINABS'],
                                          [StrToDate(LigneOrigine.DateDebutAbs),
                                           StrToDate(LigneOrigine.DateFinAbs )],
                                          TRUE);
    if (LesAbsences <> nil) then
    // Une absence pour cette date existe
    // il faut "dépointer" le mouvement d'absence
    begin
      if (LesAbsences.GetValue('PCN_IJSSSOLDEE') = 'X') then
        LesAbsences.PutValue('PCN_IJSSSOLDEE', '-');
      GAbsences.Cells[7, LesAbsences.GetValue('LIGNE')] := LesAbsences.GetValue('PCN_IJSSSOLDEE');
    end;
  end;

  if (LigneOrigine.DateDebutAbs <> Greglt.Cells[2, Ou]) or
     (LigneOrigine.DateFinAbs <>  Greglt.Cells[3, Ou]) or
// d PT2-2
     (LigneOrigine.MtRegle <>  Greglt.Cells[1, Ou]) or
     (LigneOrigine.MtIjssBrutes <>  Greglt.Cells[7, Ou]) or
// f PT2-2
     (LigneOrigine.NbJIjss <> Greglt.Cells[5, Ou]) then
  // Modif des dates d'absence ou nbre jours IJSS
  begin
    // Recherche s'il existe un mouvement pour ces dates d'absence (après modification)
    LesAbsences := TOB_Absences.FindFirst(['PCN_DATEDEBUTABS','PCN_DATEFINABS'],
                                          [StrToDate(GReglt.Cells[2,Ou]),
                                           StrToDate(GReglt.Cells[3,Ou])],
                                          TRUE);
    if (LesAbsences <> nil) then
    // Une absence pour cette date existe
    // il faut "pointer" le mouvement d'absence
    begin
      if (GReglt.Cells[5,Ou] = LesAbsences.GetValue('PCN_NBJIJSS')) then
       // le nbre de jours d'IJSS réglés = nbre de jours IJSS prévus
       // On solde la ligne d'absence
      begin
        LesAbsences.PutValue('PCN_IJSSSOLDEE', 'X');
        GAbsences.Cells[7, LesAbsences.GetValue('LIGNE')] := LesAbsences.GetValue('PCN_IJSSSOLDEE');
        PGIInfo ('Un mouvement d''absence soldé','Pointage');
      end
      else
      begin
        reponse := PGIAsk('Un mouvement d''absence pour ces dates existe mais le#10#13'+
                         ' nombre de jours IJSS prévu diffère du nombre de jours#10#13'+
                         ' IJSS réglé.#10#13'+
                         ' Voulez-vous solder le mouvement ?','Pointage');
        if (reponse=mrYes) then
          // On solde la ligne d'absence
          LesAbsences.PutValue('PCN_IJSSSOLDEE', 'X');
          GAbsences.Cells[7, LesAbsences.GetValue('LIGNE')] := LesAbsences.GetValue('PCN_IJSSSOLDEE');
      end;
    end
    else
      PGIInfo('Aucun mouvement d''absence ne correspond à ce règlement','Pointage');

  end;
end; {fin ModifReglement}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... : 02/06/2004
Description .. : procédure BTnDelClick
Suite ........ : traite l'action du bouton Bdelete
Suite ........ : vérification du pointage des absences
Suite ........ : tenu des cumuls
Suite ........ : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGLTIJSS.BTnDelClick(Sender: TObject);
var
  reponse                               : integer;
  LesAbsences                           : TOB;
begin
  reponse := mrYes;

  if (GReglt.Cells[1,GReglt.row] = '') or
     (GReglt.Cells[2,GReglt.row] = '') or
     (GReglt.Cells[3,GReglt.row] = '') then
  // règlement non saisi
  begin
    LigneValide := True;
    GReglt.DeleteRow(GReglt.row);
    ChgLigneOrigine(GReglt.row);
    if (LigneError = GReglt.row) then
    begin
      LigneError := 0;
      BtnVal.Enabled := True;
    end;
  end
  else
  begin
    If(GReglt.Cells[9,GReglt.row] <> '') then
    begin
    reponse := PGIAsk('Ce règlement a déjà été intégré dans la paie.#10#13'+
                          ' Confirmez-vous la suppression  ?',
                          ' règlement');
    end;
    if (reponse=mrYes) then
    begin
      LesAbsences := TOB_Absences.FindFirst(['PCN_DATEDEBUTABS','PCN_DATEFINABS'],
                                              [StrToDate(GReglt.Cells[2,GReglt.row]),
                                               StrToDate(GReglt.Cells[3,GReglt.row])],
                                              TRUE);
      if (LesAbsences <> nil) and not(LigneExisteDeja) then
      // Une absence pour cette date existe
      // il faut "dépointer" le mouvement d'absence s'il était soldé
      begin
        if (LesAbsences.GetValue('PCN_IJSSSOLDEE') = 'X') then
        begin
          reponse := PGIAsk('Ce règlement correspond à une absence déjà soldée.#10#13'+
                            ' Confirmez-vous la suppression  ?',
                            ' règlement');
          if (reponse=mrYes) then
          begin
            LesAbsences.PutValue('PCN_IJSSSOLDEE', '-');
            GAbsences.Cells[7, LesAbsences.GetValue('LIGNE')] := LesAbsences.GetValue('PCN_IJSSSOLDEE');
          end;
        end;
      end
      // Pas de mouvement d'absence pour ces dates
      else
        reponse := PGIAsk('Confirmez-vous la suppression de ce règlement ?',
                          'règlement');
    end;

    if (reponse=mrYes) then
    // cumuls
    begin
      if (GReglt.Cells[5,GReglt.row] <> '') and
         (GReglt.Cells[5,GReglt.row] <> '0') then
        NbjIJSSRegles := NbjIJSSRegles - Valeur(GReglt.Cells[5,GReglt.row]);
      if (GReglt.Cells[1,GReglt.row] <> '') and
         (GReglt.Cells[1,GReglt.row] <> '0,00') then
        TotIJSSNettes := TotIJSSNettes - Valeur(GReglt.Cells[1,GReglt.row]);

      // suppression de la ligne dans la grille des règlement
      LigneExisteDeja := False;
      LigneValide := True;
      if (LigneError = GReglt.row) then
      begin
        LigneError := 0;
        BtnVal.Enabled := True;
      end;
      GReglt.DeleteRow(GReglt.row);
      ChgLigneOrigine(GReglt.row);

      // affichage cumuls
      SetControlProperty('TOTIJSS', 'Value', TotIJSSNettes);
      SetControlProperty('JOURSPREV','Value', NbjIJSSPrevus);
      SetControlProperty('JOURSREGL', 'Value', NbjIJSSRegles);
      SetControlProperty('ECART', 'Value',(NbjIJSSPrevus-NbjIJSSRegles));
    end;
  end;
end; { fin BTnDelClick}

procedure TOF_REGLTIJSS.GAbsencesDblClick(Sender: TObject);
var
  Q                                     : TQuery;

begin
  if (GAbsences.col = 7) or (GAbsences.col = 8)  then exit;     // PT1
  if (GReglt.Cells[0,GReglt.RowCount-1] <> '') and
     (GReglt.Cells[1,GReglt.RowCount-1] <> '') and
     (GReglt.Cells[2,GReglt.RowCount-1] <> '') and
     (GReglt.Cells[3,GReglt.RowCount-1] <> '') and
     (GReglt.Cells[4,GReglt.RowCount-1] <> '') and
     (GReglt.Cells[5,GReglt.RowCount-1] <> '') and
     (GReglt.Cells[6,GReglt.RowCount-1] <> '') and
     (GReglt.Cells[7,GReglt.RowCount-1] <> '') and
     (GReglt.CellValues[8,GReglt.RowCount-1] <> '') then
  begin
    GReglt.RowCount := GReglt.RowCount+1;
  end;

  ChgLigneOrigine(GReglt.RowCount-1);

  // alimentation du type d'absence
  Q := OpenSql('SELECT PMA_TYPEABS FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## PMA_MOTIFABSENCE ="'+ { PT3 }
               GAbsences.CellValues[2,GAbsences.row]+'" ', TRUE);
  if not Q.eof then
  begin
    GReglt.CellValues[8,GReglt.RowCount-1] := Q.FindField('PMA_TYPEABS').Asstring;
  end
  else
    GReglt.CellValues[8,GReglt.RowCount-1] := '';
  Ferme(Q);

  GReglt.Cells[0,GReglt.RowCount-1] := DateToStr(Date);
  GReglt.Cells[2,GReglt.RowCount-1] := GAbsences.Cells[0, GAbsences.row];
  GReglt.Cells[3,GReglt.RowCount-1] := GAbsences.Cells[1, GAbsences.row];
  GReglt.Cells[4,GReglt.RowCount-1] := GAbsences.Cells[5, GAbsences.row];
  GReglt.Cells[5,GReglt.RowCount-1] := GAbsences.Cells[6, GAbsences.row];
  GReglt.Cells[6,GReglt.RowCount-1] := '';
  GReglt.Cells[7,GReglt.RowCount-1] := '';
  GReglt.Cells[9,GReglt.RowCount-1] := '';

  GReglt.row := GReglt.RowCount-1;
  GReglt.col := 0;
  GAbsences.col := 7;

  // Positionnment sur la grille des règlement
  SetFocusControl('GREGLT');
  GReglt.GoToRow(GReglt.RowCount-1);
  Action := Creation;
end; {fin GAbsencesDblClick}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 02/06/2004
Modifié le ... : 02/06/2004
Description .. : Procédure BTnValClick
Suite ........ : MAJ de la table ABSENCESALARIE à partir de la
Suite ........ : TOB_Absences mise à jour au fur et à mesure de la saisie
Suite ........ : dees règlements.
Suite ........ : MAJ de la table REGLTIJSS à partir de la grille de saisie
Suite ........ : des règlements. Les éléments initialement affichées sont
Suite ........ : d'abord supprimer de la table.
Mots clefs ... :
*****************************************************************}
procedure TOF_REGLTIJSS.BTnValClick(Sender: TObject);
var
  Ligne                         : integer;
  LesAbsences,LesReglts         : TOB;
  st,tbencours                  : string;
begin
  // traitement du cas où on valide sans "tabuler" jusqu'au bout de ligne
  if (Action = creation) then
    if  (LigneOK(GReglt.row)) and (LigneVide = False) then
    begin
      NouveauReglement(GReglt.row);
    end;
  
  if (LigneError <> 0) then
  begin
    LastError := 1;
    PGIBox('Validation impossible', '');
  end
  else
  try
    begintrans;
    tbencours := ' ABSENCESALARIE';
//    if (TOB_Absences.IsOneModifie) then
    // on procède à la maj uniquement si la TOB_Absences a été modifiée
//    begin
      for ligne := 1 to GAbsences.RowCount-1 do
      begin
        LesAbsences := TOB_Absences.FindFirst(['LIGNE'],[IntToStr(ligne)], TRUE);
        if (LesAbsences <> nil) then
        // maj du champ PCN_IJSSSOLDEE
        begin
          LesAbsences.PutValue('PCN_IJSSSOLDEE',GAbsences.Cells[7, ligne]);
          LesAbsences.PutValue('PCN_LIBCOMPL2',GAbsences.Cells[8, ligne]); // PT1

          st := 'UPDATE ABSENCESALARIE SET PCN_IJSSSOLDEE = "';
          st:= st+LesAbsences.GetValue('PCN_IJSSSOLDEE')+'" ';
// d PT1
          st := st +',PCN_LIBCOMPL2 = "';
          st:= st+Copy(LesAbsences.GetValue('PCN_LIBCOMPL2'),1,35)+'" ';
// f PT1
          st:= st+'WHERE ';
          st:= st+'(PCN_SALARIE ="';
          st:= st+LesAbsences.GetValue('PCN_SALARIE')+'") AND ';
          st:= st+'(PCN_DATEDEBUTABS ="';
          st:= st+UsDateTime(StrToDate(LesAbsences.GetValue('PCN_DATEDEBUTABS')))+'") AND ';
          st:= st+'(PCN_DATEFINABS = "';
          st:= st+UsDateTime(StrToDate(LesAbsences.GetValue('PCN_DATEFINABS')))+'") AND ';
          st:= st+'(PCN_ORDRE =';
          st:= st+intToStr(LesAbsences.GetValue('PCN_ORDRE'))+ ')';

          ExecuteSQL(st);
        end;
      end;
//    end;

    tbencours := ' REGLTIJSS';
    if TOB_Reglt <> nil then TOB_Reglt.SetAllModifie(TRUE);
    // suppression des éléments de la table REGLTIJSS affichés au chargement
    // de la fiche
    st := 'DELETE FROM REGLTIJSS ' +
          'WHERE ' +
          '(PRI_SALARIE ="' + Matricule + '") AND ' +
          '(PRI_DATEDEBUTABS >="' +
          UsDateTime(WDateDeb) + '") AND ' +
          '(PRI_DATEDEBUTABS <="' +
          UsDateTime(WDateFin) + '") AND ' +
          '(PRI_DATEFINABS >= "' +
          UsDateTime(WDateDeb) + '") AND ' +
          '(PRI_DATEFINABS <= "' +
          UsDateTime(WDateFin) + '") ';
    ExecuteSQL(st);

    FreeAndNil(TOB_Reglt);

    TOB_Reglt := TOB.Create('Les lignes de règlement', nil, -1);
    // Création de la TOB_Reglt à partir des lignes de la grille
    for ligne := 1 to GReglt.RowCount-1 do
    begin
      if (LigneOk(ligne) and (LigneVide = False)) then
      begin
        LesReglts := TOB.create('REGLTIJSS', TOB_Reglt, -1);
        LesReglts.PutValue('PRI_SALARIE', Matricule);
        LesReglts.PutValue('PRI_DATEDEBUTABS',
                           StrToDate(GReglt.Cells[2,ligne]));
        LesReglts.PutValue('PRI_DATEFINABS',
                           StrToDate(GReglt.Cells[3,ligne]));
        LesReglts.PutValue('PRI_DATEREGLT',
                           StrToDate(GReglt.Cells[0,ligne]));
        LesReglts.PutValue('PRI_NBJCARENCE',Valeur(GReglt.Cells[4,ligne]));
        LesReglts.PutValue('PRI_NBJIJSS',Valeur(GReglt.Cells[5,ligne]));
        LesReglts.PutValue('PRI_MTJOURNALIER', Valeur(GReglt.Cells[6,ligne]));
        LesReglts.PutValue('PRI_MTIJSSBRUTES', Valeur(GReglt.Cells[7,ligne]));
        LesReglts.PutValue('PRI_MTREGLE', Valeur(GReglt.Cells[1,ligne]));
        LesReglts.PutValue('PRI_TYPEABS',GReglt.CellValues[8,ligne]);
        if (GReglt.Cells[9,ligne] = '') then
          LesReglts.PutValue('PRI_DATEINTEGR',IDate1900)
        else
          LesReglts.PutValue('PRI_DATEINTEGR',StrToDate(GReglt.Cells[9,ligne]));
      end;
    end;
    TOB_Reglt.InsertDB(nil, FALSE);

    Committrans;

  except
    Rollback;
    PGIBox('! Erreur maj table'+ tbencours, '');
  end;
end;

// d PT5
procedure TOF_REGLTIJSS.BTnChercheClick (Sender: TObject);
begin
  Matricule := getControlText('PCN_SALARIE');
  WDateDeb  := StrToDate(getControlText('ABSDU'));
  WDateFin  := StrToDate(getControlText('ABSAU'));

  // Chargement de la grille des absences
  ChargementGAbsences(FALSE);

  // Chargement de la grille des règlements
  ChargementGReglt();

  // Positionnment sur la grille des règlements
  SetFocusControl('GREGLT');
  GReglt.GoToRow(GReglt.RowCount-1);

  // Affichage des différents cumuils
  SetControlProperty('TOTIJSS', 'Value', TotIJSSNettes);
  SetControlProperty('JOURSPREV', 'Value', NbjIJSSPrevus);
  SetControlProperty('JOURSREGL', 'Value', NbjIJSSRegles);
  SetControlProperty('ECART', 'Value', (NbjIJSSPrevus-NbjIJSSRegles));

  LigneValide := TRUE;
  LigneError := 0;
  BtnVal.Enabled := True;
end;
// f PT5

Initialization
  registerclasses ( [ TOF_REGLTIJSS ] ) ;
end.
