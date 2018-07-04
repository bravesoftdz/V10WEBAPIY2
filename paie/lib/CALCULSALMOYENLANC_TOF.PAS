{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 10/07/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CALCULSALMOYEN ()
                 Saisie des paramètres de calcul
Mots clefs ... : TOF;CALCULSALMOYENLANC
*****************************************************************
PT1     :    19/04/2007  NA  choix du type de calcul du salaire (avec comparaison des 3 derniers
                             mois ou salaire de la période}
Unit CALCULSALMOYENLANC_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
{$ENDIF}
     sysutils,
     HStatus,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     Vierge,
     HTB97,
     Utob,
     ed_tools,
     Utobdebug,
     pgcalculsalaire,
     HQry,
     UTOF ;

Type
  TOF_CALCULSALMOYENLANC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
     TheMul: TQUERY; // Query recuperee du mul
     Lewhere : string;
     procedure calculsalaire(Sender : Tobject);
  end ;

Implementation

procedure TOF_CALCULSALMOYENLANC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSALMOYENLANC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSALMOYENLANC.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSALMOYENLANC.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 11/07/2006
Modifié le ... :   /  /    
Description .. : ON ARGUMENT
Mots clefs ... : 
*****************************************************************}
procedure TOF_CALCULSALMOYENLANC.OnArgument (S : String ) ;
var
Btn     : TToolBarButton97;
Anneeprec : integer;
AnneeA, MoisM, jour : word;
jourj, zdate, debexerc, finexerc : Tdatetime;
codecalsal, stmode, stdatecal, stdebexerc, stfinexerc : string;

begin
  Inherited ;
  stmode := readtokenst(S);
  if stmode = 'C' then
  begin
  // Mode  lancement du calcul du salaire à partir du multicritère des salariés
  // les données du calcul sont passées en paramètre
  codecalsal := readtokenst(S);
  stdatecal := readtokenst(S);
  stdebexerc := readtokenst(S);
  stfinexerc := readtokenst(s);
  lewhere := readtokenst(s);
  SetControlText ('DATEDU', stdebexerc);
  SetControlText ('DATECALSAL',stdatecal);
  SetControlText ('DATEAU', stfinexerc);
  SetControlText ('CODECALSAL',codecalsal);
  setcontrolenabled('CODECALSAL', false);
  setcontrolenabled('DATECALSAL', false);
  setcontrolenabled('DATEDU', false);
  setcontrolenabled('DATEAU', false);
  end
  else
  begin
  // le calcul est lancé à partir du mul des salariés non calculés,
  // les données du calcul doivent être saisies
  lewhere := readtokenst(s);
  Jourj := Date;
  decodedate(Jourj,AnneeA, MoisM, Jour);
  Anneeprec := AnneeA - 1;
  zdate := Encodedate(anneeprec, moisM, Jour);
  debexerc := debutannee(zdate);
  finexerc := finannee(Zdate);
  SetControlText ('DATEDU',DateTimeToStr(debexerc));
  SetControlText ('DATECALSAL', DateTimeToStr(finexerc));
  SetControlText ('DATEAU', DateTimeToStr(finexerc));
  end;
    if TFVierge(Ecran) <> nil then
    begin
    {$IFDEF EAGLCLIENT}
    //TheMul := THQuery(TFVierge(Ecran).FMULQ).TQ;
    TheMul := THQuery(TFVierge(Ecran).FMULQ).TQ;
    {$ELSE}
    TheMul := TFVierge(Ecran).FMULQ;
    {$ENDIF}
    end;


   Btn := TToolBarButton97(GetControl('B_VALIDER'));
  if btn<>nil then btn.OnClick := calculsalaire;
end ;

procedure TOF_CALCULSALMOYENLANC.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSALMOYENLANC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSALMOYENLANC.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 11/07/2006
Modifié le ... :   /  /    
Description .. : CALCUL DU SALAIRE MOYEN POUR LES SALARIES 
Suite ........ : SELECTIONNES
Mots clefs ... : 
*****************************************************************}
procedure TOF_CALCULSALMOYENLANC.calculsalaire(sender : Tobject);
var
tob_erreur, TTerreur : Tob;
codecal, sal, nomsal, prenomsal : string;
datecal, debexerc, finexerc : TDatetime;
Listanomalies : Tlistbox;
wsalairecomp : Tcheckbox; // pt1
Salcompare : boolean; // pt1

begin
codecal := getcontroltext('CODECALSAL');
datecal :=  StrToDate(getcontroltext('DATECALSAL'));
Debexerc := StrToDate(getcontroltext('DATEDU'));
finexerc := StrToDate(getcontroltext('DATEAU'));

// contrôle des saisies
if codecal = '' then begin
PGIBOX('Le code calcul doit être renseigné.', Ecran.caption);
setfocuscontrol('CODECALSAL');
exit;
end
else
begin
// contrôle qu'une méthode de calcul a été saisie
  if existesql('select psm_codecalsal from METHCALCULSALMOY where psm_codecalsal = "'+codecal+'"') = false then
  begin
  PGIBOX('Aucune méthode de calcul n''a été renseignée pour ce code.', Ecran.caption);
  setfocuscontrol('CODECALSAL');
  exit;
  end;
end;


if (finexerc < debexerc) then
begin
PGIBOX('Période invalide', Ecran.caption);
setfocuscontrol('DATEDU');
exit;
end;

if ((debexerc > datecal) or (finexerc < datecal)) then
begin
PGIBOX('La date de calcul n''est pas comprise dans la période.');
setfocuscontrol('DATECALSAL');
exit;
end;

// deb pt1
wsalairecomp := TcheckBox(getcontrol('SALAIRECOMP'));
if wsalairecomp.checked then salcompare := true else salcompare := false;
// fin pt1


 If TheMul = nil then exit;

//TheMul.First;
InitMoveProgressForm (NIL,'Traitement de calcul en cours', 'Veuillez patienter SVP ...',theMul.RecordCount,FALSE,TRUE);



// pt1 calculsalmoyen(tob_sal, codecal,datecal, debexerc, finexerc, tob_erreur);
calculsalmoyen(lewhere, codecal,datecal, debexerc, finexerc,salcompare, tob_erreur);   // pt1

// affichage des erreurs sur la lISTBOX
listanomalies := TListbox(Getcontrol('LBANOMALIES'));
listanomalies.Items.clear;
   TTerreur := TOB_erreur.FindFirst([''], [''], TRUE);
   while TTerreur <> nil do
   begin
   sal := TTerreur.getvalue('PSA_SALARIE');
   nomsal := TTerreur.getvalue('PSA_LIBELLE');
   prenomsal := TTerreur.getvalue('PSA_PRENOM');
   listanomalies.items.Add('le salarié ' + sal + ' ' + nomsal + ' ' + prenomsal+ ' n''a pas de salaire sur la période');
   TTerreur := TOB_erreur.Findnext([''], [''], TRUE);
   end;
   
//tobdebug(tob_erreur);

freeandnil(tob_erreur);
FiniMoveProgressForm;
PgiInfo('Traitement terminé.','Calcul des salaires moyens');
end;


Initialization
  registerclasses ( [ TOF_CALCULSALMOYENLANC ] ) ;
end.
