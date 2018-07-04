{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 20/03/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CALCULSALMOY () Lancement du calcul du salaire moyen
Mots clefs ... : TOF;CALCULSALMOY
*****************************************************************}
Unit CALCULSALMOY_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     HDB,
     fe_main,
{$else}
     eMul,
     uTob,
     maineagl,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97,
     P5Def,
     utof,
     Hqry,
     HStatus,
     AglInit;


Type
  TOF_CALCULSALMOY = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
 private
   {$IFNDEF EAGLCLIENT}
    Liste: THDBGrid;
{$ELSE}
    Liste: THGrid;
{$ENDIF}

    codecal : string;
    datecal, debexerc, finexerc : Tdatetime;
    lancecalcul : boolean;
    procedure calculsalaire(Sender : Tobject);
    end;

Implementation

procedure TOF_CALCULSALMOY.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSALMOY.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSALMOY.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 11/07/2006
Modifié le ... :   /  /    
Description .. : chargement
Mots clefs ... : 
*****************************************************************}
procedure TOF_CALCULSALMOY.OnLoad ;
var
datesaisi : string;

begin
  Inherited ;
  datesaisi := getcontrolText('DATECALCUL');
  Datecal := StrToDate(Datesaisi);
  if Lancecalcul then exit;
  SetControlText('XX_WHERE','(PSA_DATESORTIE >="'+usdatetime(datecal)+'" or'+
  ' PSA_DATESORTIE <="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) and PSA_DATEENTREE <= "'+
  usdatetime(datecal)+'"');
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 11/07/2006
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_CALCULSALMOY.OnArgument (S : String ) ;
var
Btn : TToolBarButton97;
Anneeprec, num : integer;
AnneeA, MoisM, jour : word;
jourj, zdate : Tdatetime;
begin
  Inherited ;
  Jourj := Date;
  decodedate(Jourj,AnneeA, MoisM, Jour);
  Anneeprec := AnneeA - 1;
  zdate := Encodedate(anneeprec, moisM, Jour);
  debexerc := debutannee(zdate);
  finexerc := finannee(Zdate);
  SetControlText ('DATEDEBUT',DateTimeToStr(debexerc));
  SetControlText ('DATEFIN', DateTimeToStr(finexerc));
  SetControlText ('DATECALCUL', DateTimeToStr(finexerc));

  for Num := 1 to 4 do
  begin
  VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));

  Btn := TToolBarButton97(GetControl('Bouvrir'));
  if btn <> nil then Btn.Onclick := Calculsalaire;

  end ;

procedure TOF_CALCULSALMOY.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSALMOY.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSALMOY.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 11/07/2006
Modifié le ... :   /  /    
Description .. : calcul salaire
Mots clefs ... : 
*****************************************************************}
procedure TOF_CALCULSALMOY.Calculsalaire (sender : Tobject);
var
Suite, i : integer;
dt, salarie , st, param, stmode, lewhere : string;
Q_mul : THQuery;

begin

stmode := 'C';
{$IFNDEF EAGLCLIENT}
Liste:= THDBGrid(GetControl('FListe'));
{$ELSE}
Liste:= THGrid(GetControl('FListe'));
{$ENDIF}

codecal := getcontroltext('CODECALSAL');
dt:= getcontroltext('DATECALCUL');
datecal :=  StrToDate(getcontroltext('DATECALCUL'));
Debexerc := StrToDate(getcontroltext('DATEDEBUT'));
finexerc := StrToDate(getcontroltext('DATEFIN'));

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
PGIBOX('Période invalide.', Ecran.caption);
setfocuscontrol('DATEDEBUT');
exit;
end;

if ((debexerc > datecal) or (finexerc < datecal)) then
begin
PGIBOX('La date de calcul n''est pas comprise dans la période.');
setfocuscontrol('DATECALCUL');
exit;
end;



if Liste <> nil then
   begin   // D1
   if (Liste.NbSelected = 0) and (not Liste.AllSelected) then
      begin
      PGIBOX('Aucun salarié n''est sélectionné.', Ecran.Caption);
      exit;
      end;
    suite := Pgiask('Calcul du salaire moyen pour le code calcul '+Codecal+' '+
    'à la date du '+dt+' #13#10 pour les salariés sélectionnés. '+'Voulez vous poursuivre?' ,Ecran.Caption);
    if (suite <> mrYes) then exit;

// lecture de la liste des salariés

    //{$IFDEF EAGLCLIENT}
   // Q_mul := TOB(Ecran.FindComponent('Q'));
  //  {$ELSE}
   Q_mul := THQuery(Ecran.FindComponent('Q'));
    //{$ENDIF}

      St := '';
      if (Liste.AllSelected = false) then

     // lecture de chaque salarié sélectionné
      begin

            for i := 0 to Liste.NbSelected - 1 do
            begin
           {$IFDEF EAGLCLIENT}
            TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
           {$ENDIF}
            Liste.GotoLeBOOKMARK(i);
            salarie := Q_mul.findfield('PSA_SALARIE').asstring;
            St := St + ' PSA_SALARIE="' + Salarie + '" OR';
            end;
       Liste.ClearSelected;
     end; 

  // Calcul du salaire


   if St <> '' then
   begin
      St := Copy(St, 1, Length(st) - 2)  ;
      SetControlText('XX_WHERE',  St);
      Lancecalcul := True;
      TFMul(Ecran).BCherche.Click;
      Lancecalcul := False;
   end;

    { Récupération de la Query pour traitement dans la fiche vierge }
{$IFDEF EAGLCLIENT}
     if (TFMul(Ecran).bSelectAll.Down) then
         TFMul(Ecran).Fetchlestous;

  TheMulQ := TOB(Ecran.FindComponent('Q'));
{$ELSE}
  TheMulQ := THQuery(Ecran.FindComponent('Q'));
{$ENDIF}
  {Ouverture de la fiche}
  LeWhere := RecupWhereCritere(TFMul(Ecran).Pages);
  param := stmode + ';' +  codecal + ';' + getcontroltext('DATECALCUL') +  ';' +
  getcontroltext('DATEDEBUT') + ';' + getcontroltext('DATEFIN')+ ';' + lewhere;
  AglLanceFiche('PAY', 'CALCULSALMOYEN', '', '', param);
  TheMulQ := nil;

  TFMul(Ecran).BCherche.Click;
          

  end; // E1
end;



Initialization
  registerclasses ( [ TOF_CALCULSALMOY ] ) ;
end.
