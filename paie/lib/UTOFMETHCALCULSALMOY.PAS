{***********UNITE*************************************************
Auteur  ...... :  NA
Créé le ...... : 06/09/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOFMETHCALCULSALMOY ()
Mots clefs ... : TOF;UTOFMETHCALCULSALMOY
(remplace METHCALCULSALMOY_TOM)
*****************************************************************}
Unit UTOFMETHCALCULSALMOY ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     Graphics,
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     Fiche,
{$else}
     eFiche,
     eMul, 
{$ENDIF}
     grids,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     PgOutils,
     P5Util,
     UTOF,
     HTB97,
     UTob  ;

Type
  TOF_METHCALCULSALMOY = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure Gestionpred;


private
    Grille: THGrid;
    LeTitre,pred,  nodos, codecalsal, libelle : string;
    LectureSeule, CEG, STD, DOS, MAJOK, AmettreAjour : boolean;
    procedure GrilleDeleteLigne(sender: TObject);
    procedure validesaisie(sender : TObject);
    procedure Alimgrille;
    procedure Initgrille;
    procedure GrillePostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure Trtnature(sender: TObject);
    procedure trtrubcum(Sender : TObject);
    procedure Validerclick(sender: TObject);
    end ;


Implementation

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/09/2007
Modifié le ... :   /  /    
Description .. : Nouvel enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOF_METHCALCULSALMOY.OnNew ;
begin
  Inherited ;

setcontroltext('CODECALSAL', Codecalsal);
setcontroltext('LIBELLECAL', Libelle);
setcontroltext('PREDEFINI', Pred);
setcontroltext('NODOSSIER', Nodos);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/09/2007
Modifié le ... :   /  /    
Description .. : Suppression
Mots clefs ... : 
*****************************************************************}
procedure TOF_METHCALCULSALMOY.OnDelete ;
var
suite : integer;
begin
  Inherited ;
    suite := Pgiask('Confirmez-vous la suppression de l''enregistrement? ' ,Ecran.Caption);
    if (suite = mrYes) then
    begin
      ExecuteSQL('DELETE FROM METHCALCULSALMOY WHERE PSM_PREDEFINI="' + pred +'"'+
      ' AND PSM_NODOSSIER ="' +nodos + '" AND PSM_CODECALSAL="' + codecalsal+'"');
      Alimgrille;
    end;
end ;

procedure TOF_METHCALCULSALMOY.Onupdate ;
begin
  Inherited ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/09/2007
Modifié le ... :   /  /    
Description .. : Validation 
Mots clefs ... : 
*****************************************************************}
procedure TOF_METHCALCULSALMOY.ValiderClick(Sender:TObject);
var
  Tob_Methode, T1: TOB;
  naturecal, rubcum, basemont, sens, sql1, proratise : string;  // pt1
  i, suite : Integer;
  saisieproratise : string;
begin
  Inherited ;
    // si ce code calcul a déjà été utilisé: message d'alerte
  Sql1 := 'Select PSY_CODECALSAL from SALAIREMOYEN where PSY_CODECALSAL = "'+CODECALSAL+'"';
  if ExisteSql(Sql1) then
  begin
    suite := Pgiask('Attention ! Ce code calcul a déjà été utilisé.#13#10'+
   ' Les salaires moyens calculés risquent d''être incohérents avec cette méthode de calcul.#13#10'+
   ' Voulez vous poursuivre ?' ,Ecran.Caption);
  if (suite <> mrYes) then exit;
  end;
    // Mise à jour
  if Grille.rowcount - 1  >= 1 then
  begin
    Tob_Methode := TOB.CREATE('_Methodecalcul', nil, -1);
    for i := 1 to grille.rowcount - 1 do
    begin
      naturecal := Grille.Cellvalues[0, i];

      if naturecal <> '' then
      begin
        basemont := Grille.Cellvalues[1, i];
        rubcum := Grille.Cellvalues[2, i];
        sens   := Grille.Cellvalues[4, i];
        // pt1
        proratise := Grille.Cellvalues[6, i];
        if proratise = 'OUI' then
        saisieproratise := 'X'
        else saisieproratise := '-';
        // fin pt1

        T1 := TOB.CREATE('METHCALCULSALMOY', Tob_Methode, -1);
        T1.PutValue('PSM_PREDEFINI', PRED);
        T1.PutValue('PSM_NODOSSIER', NODOS);
        T1.PutValue('PSM_CODECALSAL', CODECALSAL);
        T1.PutValue('PSM_NATURECAL', naturecal);
        T1.PutValue('PSM_BASEMONT', basemont);
        T1.PutValue('PSM_RUBRCUM', rubcum);
        T1.PutValue('PSM_SENSCALC', sens);
        T1.PutValue('PSM_PRORATISE', saisieproratise); // pt1
      end;
    end;
     ExecuteSQL('DELETE FROM METHCALCULSALMOY WHERE PSM_PREDEFINI="' + PRED +'" '+
    ' AND PSM_NODOSSIER ="' + nodos + '" AND PSM_CODECALSAL="' + codecalsal+'"');
    Tob_Methode.InsertDB(nil, false);
    MAJOK := true;
    Setcontrolvisible('BDELETE', true);
    freeandnil(Tob_Methode);
   end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/09/2007
Modifié le ... :   /  /    
Description .. : Chargement des données
Mots clefs ... : 
*****************************************************************}
procedure TOF_METHCALCULSALMOY.OnLoad ;
begin
  Inherited ;
  if Grille <> nil then
initgrille;    // Initialise la grille
Alimgrille;    // Alimente la grille avec les données déjà saisies

setcontroltext('CODECALSAL', Codecalsal);
setcontrolenabled('CODECALSAL', false);
setcontroltext('LIBELLECAL', Libelle);
setcontrolenabled('LIBELLECAL', false);
setCONTROLTEXT('PREDEFINI', Pred);
setcontrolenabled('PREDEFINI', false);
setcontroltext('NODOSSIER', Nodos);
setcontroltext('NATURECAL','');
setcontroltext('BASEMONT','');
setcontroltext('SENSCALC','');
setcontroltext('RUBRCUM','');
setcontroltext('LIBELLE','');


end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/09/2007
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_METHCALCULSALMOY.OnArgument (S : String ) ;
var
  Btn: TToolBarButton97;
  COMBONATCAL : THValComboBox;
  THrubcum : THedit;
begin
  Inherited ;
   MajOK := false;
   AMettreAjour := false;

   Grille := THGrid(GetControl('DETAILMINI'));
  if Grille <> nil then
  begin
    Btn := TToolBarButton97(GetControl('B_SUPPRESSION'));
    if Btn <> nil then Btn.OnClick := GrilleDeleteLigne;
    Btn := TToolBarButton97(GetControl('B_OKSAISIE'));
    if Btn <> nil then Btn.OnClick := Validesaisie;
     Grille.PostDrawCell := GrillePostDrawCell;
  end;

  Btn := TToolBarButton97(GetControl('VALIDE'));
  if Btn <> nil then Btn.OnClick := ValiderClick;

  LeTitre := 'Méthode de calcul';
  Pred   := ReadTokenSt(S);
  Nodos := ReadTokenSt(S);
  Codecalsal := ReadTokenSt(S);
  libelle := ReadTokenSt(S);

  AccesPredefini('TOUS', CEG, STD, DOS);

  COMBONATCAL := THValComboBox(GetControl('naturecal'));
  Thrubcum := THEdit(GetControl('RUBRCUM'));
  If ComboNATCAL <> Nil then ComboNATCAL.OnExit := Trtnature;
  if Thrubcum <> NIL then thrubcum.OnExit := trtrubcum;
  if Thrubcum <> NIL then thrubcum.Onclick := trtrubcum;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/09/2007
Modifié le ... :   /  /
Description .. : Saisie rubrique ou cumul
Mots clefs ... :
*****************************************************************}
procedure TOF_METHCALCULSALMOY.trtrubcum;
var
rubcumval, lib : string;

begin
lib := '';
rubcumval := THEdit(GetControl('RUBRCUM')).text;

if rubcumval <> '' then
begin

setcontroltext('LIBELLE','');
 if THValComboBox(GetControl('naturecal')).value = 'REM'  then
 begin                                               // si remunération
lib := rechdom('PGREMUNERATION', rubcumval, false);
   if (lib = '') or (lib = 'Error') then
   begin
   PGIBOX('Cette rubrique n''existe pas.', Letitre);
   setfocuscontrol('RUBRCUM');
   end
   else
   setcontrolcaption('LIBELLE', lib);
 end
 else
 begin                                              // si cumul
lib := rechdom('PGCUMULPAIE', rubcumval, false);
   if (lib = '') or (lib = 'Error') then
   begin
   PGIBOX('Ce cumul n''existe pas.', Letitre);
   setfocuscontrol('RUBRCUM');
   end
   else
   setcontrolcaption('LIBELLE', lib);
 end;
end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/09/2007
Modifié le ... :   /  /
Description .. : Click sur nature calcul
Mots clefs ... :
*****************************************************************}
procedure TOF_METHCALCULSALMOY.Trtnature;
begin
  setcontroltext('BASEMONT','');
  setcontroltext('SENSCALC','');
  setcontroltext('RUBRCUM','');
  setcontroltext('LIBELLE','');

  if THValComboBox(GetControl('naturecal')).value = 'CUM' then // si cumul
  begin
  setcontrolenabled('BASEMONT', false);
  setcontrolproperty('RUBRCUM','Datatype','PGCUMULPAIE');
  setfocuscontrol('RUBRCUM');
  end
  else  // si rem
  begin
  setfocuscontrol('BASEMONT');
  setcontrolenabled('BASEMONT', true);
  setfocuscontrol('BASEMONT');
  setcontrolproperty('RUBRCUM','Datatype', 'PGREMUNERATION');
  end;
  end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/09/2007
Modifié le ... :   /  /    
Description .. : Validation de la saisie pour affichage dans la grille
Mots clefs ... : 
*****************************************************************}
procedure TOF_METHCALCULSALMOY.validesaisie;
var
arow, i : integer;
naturecalval, basemontval, rubrcumval,senscalcval, lib, proratise : string;
saisieok : boolean;
begin
saisieok := true;
naturecalval := THValComboBox(GetControl('naturecal')).value ;
basemontval :=  THValComboBox(GetControl('basemont')).value ;
rubrcumval :=   THEdit(GetControl('RUBRCUM')).text;
senscalcval :=  THValComboBox(GetControl('Senscalc')).value;
proratise :=  THValComboBox(GetControl('PRORATISE')).value;   // pt1

// contrôle de la saisie

if naturecalval = '' then
begin
PGIBOX('La nature du calcul doit être renseignée.', Letitre);
setfocuscontrol('NATURECAL');
exit;
end;

if (naturecalval = 'REM') and (basemontval = '') then
begin
PGIBOX('La base ou le montant doit être renseigné.', Letitre);
setfocuscontrol('BASEMONT');
exit;
end;

if rubrcumval = '' then
begin
if naturecalval = 'REM' then PGIBOX('La rubrique doit être renseignée.', Letitre) else
PGIBOX('Le cumul doit être renseigné.', Letitre);
setfocuscontrol('RUBRCUM');
exit;
end;

lib :=THlabel(GetControl('LIBELLE')).caption;
if lib= '' then
begin
if naturecalval = 'REM' then PGIBOX('La rubrique n''existe pas.', Letitre) else
PGIBOX('Le cumul n''existe pas.', Letitre);
setfocuscontrol('RUBRCUM');
exit;
end;

if senscalcval = '' then
begin
PGIBOX('Le sens doit être renseigné.', Letitre);
setfocuscontrol('SENSCALC');
exit;
end;

//pt1
if (proratise <> 'OUI') and (proratise <> 'NON') then
begin
PGIBOX('Proratisé : OUI ou NON.',Letitre);
setfocuscontrol('PRORATISE');
exit;
end;
// fin pt1

// contrôle que la ligne saisie n'esxite pas dans la grille

for i := 1 to Grille.rowcount - 1 do
begin
if (Grille.cellvalues[2,i] = rubrcumval) then
begin
 if (naturecalval = 'REM') then PGIBOX('Cette rubrique a déjà été saisie.', Letitre) else
 PGIBOX('Ce cumul a déjà été saisi.', Letitre);
 setfocuscontrol('RUBRCUM');
 saisieok := false;
 break;
 end;
end;

if (saisieok = true) then
begin

// ajout de la saisie dans la grille

if Grille.cellValues[0,1] = '' then arow := 1 else   // si grille vide
begin   // ajout d'une ligne dans la grille
Grille.rowcount := Grille.rowcount + 1;
arow := grille.rowcount -1;
end;

Grille.cellvalues[0, arow] := naturecalval;
Grille.cellvalues[1, arow] := basemontval;
Grille.cellvalues[2, arow] := rubrcumval;
Grille.cellvalues[3, arow] := THlabel(GetControl('LIBELLE')).caption;
Grille.cellvalues[4, arow] := senscalcval;
Grille.cellvalues[6, arow] := proratise; // pt1

setcontroltext('NATURECAL','');
setcontroltext('BASEMONT','');
setcontroltext('SENSCALC','');
setcontroltext('RUBRCUM','');
setcontroltext('LIBELLE','');
setcontroltext('PRORATISE','');  // pt1
AMettreAjour := true;
setcontrolvisible('BDELETE', false);
//if not (ds.state in [dsinsert, dsedit]) then ds.edit;
end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 21/04/2006
Modifié le ... :   /  /
Description .. : SUPPRESSION DE LIGNES
Mots clefs ... : 
*****************************************************************}
procedure TOF_METHCALCULSALMOY.GrilleDeleteLigne(sender: TObject);
begin
  if Grille.Row <> 0 then Grille.deleterow(Grille.Row);
  AMettreAjour := true;
  Setcontrolvisible('BDELETE', false);
  //if not (ds.state in [dsinsert, dsedit]) then ds.edit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/09/2007
Modifié le ... :   /  /
Description .. : Alimente la grille
Mots clefs ... : 
*****************************************************************}
procedure TOF_METHCALCULSALMOY.AlimGrille;
var
  TOB_METHCAL: TOB;
  Methcalcul, naturecal, zonebl, saisieproratise: string;
  QMETHCAL: Tquery;
  i       : integer;


begin
     Methcalcul := 'SELECT PSM_NATURECAL, PSM_BASEMONT, PSM_RUBRCUM, PSM_SENSCALC, PSM_PRORATISE from METHCALCULSALMOY' +
      ' where PSM_PREDEFINI = "' + pred + '" and PSM_NODOSSIER = "' + nodos + '"' +
      ' and PSM_CODECALSAL = "' + codecalsal + '" ';
    QMETHCAL := OpenSQL(Methcalcul, True);

    zonebl := '';           
    TOB_METHCAL := TOB.CREATE('Methode calcul', nil, -1);
    TOB_METHCAL.LoadDetailDB('METHCALCULSALMOY', '', '', QMETHCAL, False);

    TOB_METHCAL.PutGridDetail(Grille, false, false, 'PSM_NATURECAL;PSM_BASEMONT;PSM_RUBRCUM;zonebl;PSM_SENSCALC;PSM_PRORATISE;zonebl', True);

    freeandnil(TOB_METHCAL);

    Ferme(QMETHCAL);
  //  Grille.ColFormats[1] := 'CB=PGTYPECHAMPCALC';

     for i := 1 to grille.rowcount - 1 do
     begin
     naturecal := Grille.CellValues[0,i];
      if naturecal <> '' then
      begin

        if Grille.CellValues[0,i] = 'REM' then
        Grille.CellValues[3,i] := Rechdom('PGREMUNERATION', Grille.CellValues[2,i],False)
        else
        Grille.CellValues[3,i] := Rechdom('PGCUMULPAIE', Grille.CellValues[2,i],False);
      //  Grille.ColFormats[1] := '';
     // pt1 : Transforme le booléen en OUI/NON
        saisieproratise := Grille.CellValues[5,i];
        if saisieproratise = 'X' then Grille.CellValues[6,i] := 'OUI' else
        Grille.CellValues[6,i] := 'NON';
     // fin pt1
      end;
     end;


end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/09/2007
Modifié le ... :   /  /    
Description .. : Initialise la grille
Mots clefs ... : 
*****************************************************************}
procedure TOF_METHCALCULSALMOY.Initgrille;
begin
   // gestion des prédéfinis
  gestionpred;

  setcontrolenabled('B_SUPPRESSION', true);
  setcontrolenabled('B_OKSAISIE', true);
  if lectureseule = true then
  begin
   setcontrolenabled('B_SUPPRESSION', false);
   setcontrolenabled('B_OKSAISIE', false);
  end;
  if lectureseule = false then
  begin
  Grille.ColEditables[0] := FALSE;
  Grille.ColEditables[2] := FALSE;
  Grille.ColEditables[4] := FALSE;
  Grille.ColEditables[1] := false;
  end;
  Grille.Elipsisbutton := false;
  Grille.ColFormats[0] := 'CB=PGNATURECALCUL';
  Grille.ColFormats[1] := 'CB=PGTYPECHAMPCALC';
  Grille.ColFormats[4] := 'CB=PGSENSCALCUL';
  Grille.ColFormats[6] := 'CB=PGOUINON'; // pt1

  Grille.Colcount := 7;   // pt1
  Grille.Colwidths[0] := 97;
  Grille.Colwidths[1] := 106;
  Grille.Colwidths[2] := 100;
  Grille.Colwidths[3] := 165;
  Grille.Colwidths[4] := 60;
  Grille.Colwidths[6] := 60;   // pt1
  Grille.colwidths[5] := -1; // pt1
end;

{
{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 24/04/2006
Modifié le ... :   /  /
Description .. : Gestion des prédéfinis
Mots clefs ... : 
*****************************************************************}
procedure TOF_METHCALCULSALMOY.gestionpred ;
begin
// Gestion des prédéfinis
  LectureSeule := FALSE;

  if pred = 'CEG' then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
  end;

  if pred = 'STD' then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
  end;

  if pred = 'DOS' then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), False);
  end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/09/2007
Modifié le ... :   /  /    
Description .. : On close
Mots clefs ... : 
*****************************************************************}
procedure TOF_METHCALCULSALMOY.OnClose ;
var
suite : integer;
begin
  Inherited ;
    if (MAJOK = false) and (AMettreAjour = true) then
    begin
       suite := Pgiask('Voulez-vous enregistrer les modifications ? ' ,Ecran.Caption);
      if (suite = mrYes) then Validerclick(nil);
    end;
end ;

procedure TOF_METHCALCULSALMOY.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_METHCALCULSALMOY.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/09/2007
Modifié le ... :   /  /
Description .. : Procédure dessinant les cellules
Mots clefs ... :
*****************************************************************}
procedure TOF_METHCALCULSALMOY.GrillePostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
begin
  if (ARow <> 0) then GridGriseCell(Grille, Acol, Arow, Canvas);
end;


Initialization
  registerclasses ( [ TOF_METHCALCULSALMOY ] ) ; 
end.
