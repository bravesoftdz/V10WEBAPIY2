{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 08/06/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CALCULSIMULPARAM ()
                 Calcul de la simulation : Saisie des paramètres de la simulation
Mots clefs ... : TOF;CALCULSIMULPARAM
*****************************************************************}
Unit CALCULSIMULPARAM_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     FE_Main,
{$else}
     MaineAGL,
     uTob, 
{$ENDIF}
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Vierge,
     lookup,
     HMsgBox,
     HTB97,
     Entpaie,
     Pgpopuloutils,
     UTOF ;

Type
  TOF_CALCULSIMULPARAM = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    mode , codesimd : string;
    procedure Calculsimul(sender: TObject);
    procedure DatesimulExit(sender: TObject);
    procedure DateCalculExit(sender: TObject);
    procedure codecalculexit(sender : TObject);
    procedure Codecalculelipsisclick(sender: TObject);
    procedure recherchetauxactu;
   end ;

Implementation

procedure TOF_CALCULSIMULPARAM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSIMULPARAM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSIMULPARAM.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSIMULPARAM.OnLoad ;
var
popvalide : boolean;
pred : string;
begin
  Inherited ;
  if mode = 'C' then
  begin
     pred := GetPredefiniPopulation('IDR');
     recherchetauxactu;
     popvalide := canusepopulation('IDR');
     if not popvalide then PGIINFO('Attention : les populations IDR (prédéfini '+pred+') ne sont pas valides.');
  end;
  setfocuscontrol('DATESIMUL');

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/06/2006
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_CALCULSIMULPARAM.OnArgument (S: string);
var
Btn : TToolBarButton97;
DD, DDCAL : THedit;
Anneeprec :integer;
AnneeA, MoisM, jour : word;
jourj, zdate, dtcalcul, dtsimul : Tdatetime;
codesim : string;
QQ: Tquery;
codecalcul : THEdit;

begin
  Inherited ;
  mode := readtokenst(S);
  codesim := readtokenst(S);
  codesimd := codesim;

  if mode = 'D' then
  Ecran.caption := 'Duplication de la simulation' + ' ' + codesim
  else
  if mode = 'M' then Ecran.Caption := 'Modification de la simulation' + ' ' + codesim;
  
  if mode = 'C'  then begin                         // si création
  Jourj := Date;
  decodedate(Jourj,AnneeA, MoisM, Jour);
 // Date calcul = date fin année précédente //
 // Anneeprec := AnneeA - 1;
 // zdate := Encodedate(anneeprec, moisM, Jour);
 // dtcalcul := finannee(Zdate);
 // SetControlText ('DATECALC',DateTimeToStr(dtcalcul));

  // date de simulation = date fin de mois précédente //
  dtsimul := FinDeMois(Plusmois(jourj, -1));
  SetControlText ('DATESIMUL',DateTimeToStr(dtsimul));
  end
  else                                            // si duplication ou modification
  begin
    // Recherche des données de la simulation à dupliquer ou modifier
   QQ := opensql('select * from SIMULATION where PSL_CODESIMUL = "'+codesim+'"', true);
   if not QQ.EOF then
   begin
   setcontroltext('DATESIMUL', QQ.findfield('PSL_DATESIMUL').asstring);
   setcontroltext('CODESIMUL', QQ.findfield('PSL_CODESIMUL').asstring);
   setcontroltext('LIBELLESIMUL', QQ.findfield('PSL_LIBELLE').asstring);
   setcontroltext('TYPEDEPART', QQ.findfield('PSL_TYPEDEPART').asstring);
   setcontroltext('METHCALCUL', QQ.findfield('PSL_METHCALCUL').asstring);
   setcontroltext('TAUXACTUAL', QQ.findfield('PSL_TAUXACTUAL').asstring);
   setcontroltext('CODECALCUL', QQ.findfield('PSL_CODECALSAL').asstring);
   setcontroltext('DATECALC', QQ.findfield('PSL_DATECALC').asstring);
   end;
   ferme(QQ);

   if mode = 'M' then begin
   setcontrolenabled('DATESIMUL', false);
   setcontrolenabled('CODESIMUL', false);
   end;
   end;

  DD:= THedit(Getcontrol('DATESIMUL'));
  DDCAL := THedit(Getcontrol('DATECALC'));
  if DD <> nil then DD.OnClick := DatesimulExit;
  if DD <> nil then DD.onexit := Datesimulexit;


  codecalcul := THEdit(GetControl('CODECALCUL'));
  If codecalcul <> Nil Then codecalcul.OnElipsisClick := codecalculElipsisClick;
  if codecalcul <> Nil Then codecalcul.Onexit := codecalculexit;

  if DDCAL <> nil then DDCAL.OnClick := DateCalculExit;
  if DDCAL <> nil then DDCAL.onExit := DateCalculExit;
  
  Btn := TToolBarButton97(GetControl('BDACCORD'));
  if btn <> nil then  Btn.Onclick := calculsimul;

end ;


procedure TOF_CALCULSIMULPARAM.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSIMULPARAM.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSIMULPARAM.OnCancel () ;
begin
  Inherited ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/06/2006
Modifié le ... :   /  /
Description .. : Contrôle date simulation si click sur date
Mots clefs ... :
*****************************************************************}
procedure TOF_CALCULSIMULPARAM.DatesimulExit(Sender: TObject);
var
  DD: THEdit;
begin
  DD := THEdit(GetControl('DATESIMUL'));

  if (DD <> nil) and (mode = 'C') then recherchetauxactu;
 
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/06/2006
Modifié le ... :   /  /
Description .. : Contrôle date calcul salaire si click sur date
Mots clefs ... :
*****************************************************************}
procedure TOF_CALCULSIMULPARAM.DateCalculExit(Sender: TObject);
var
  DD: THEdit;
  begin
  DD := THEdit(GetControl('DATECALC'));

  if (DD <> nil)  then
  begin
    if not (isvaliddate(DD.text)) then
    begin
      PGIBox('Date de calcul des salaires incorrecte.', Ecran.Caption);
      SetFocusControl('DATECALC');
      exit;
    end;
  end;
  end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/06/2006
Modifié le ... :   /  /
Description .. : Calcul simulation
Mots clefs ... :
*****************************************************************}
procedure TOF_CALCULSIMULPARAM.Calculsimul ;
var
dtsimul,dtcalcul, codesimul, libellesimul, typedepart, methcalcul, codecalcul,txactu, St, sql1, sql2, retour : string;
datecalcul : Tdatetime;
tauxactual : double;
ctlok : boolean;

begin
  codesimul:= getcontroltext('CODESIMUL');
  dtsimul:= getcontroltext('DATESIMUL');
  libellesimul := getcontroltext('LIBELLESIMUL');

  typedepart := THValcomboBox(getcontrol('TYPEDEPART')).value;
  dtcalcul := getcontroltext('DATECALC');
  methcalcul := THValcomboBox(getcontrol('METHCALCUL')).value;

  codecalcul := getcontroltext('CODECALCUL');
  tauxactual:= strtofloat(getcontroltext('TAUXACTUAL'));

  datecalcul :=  StrToDate(getcontroltext('DATECALC'));
  txactu := floattoStr(tauxactual);

  ctlok := true;
  // Controle de la saisie

  if codesimul = '' then begin
  ctlok := false;
  PGIBOX('Le code simulation doit être renseigné.', Ecran.caption);
  setfocuscontrol('CODESIMUL');
  end;

  // controle que ce code simulation n'existe pas déjà si non modification
  if mode <> 'M' then begin

  Sql2 := 'Select PSL_CODESIMUL from SIMULATION where PSL_CODESIMUL = "'+Codesimul+'"';
  if (ExisteSql(Sql2)  and ctlok) then
  begin
  ctlok := false;
  PGIBOX('Ce code simulation existe déjà.', Ecran.caption);
  setfocuscontrol('CODESIMUL');
  end;
  end;

  if ((libellesimul = '') and ctlok ) then begin
  ctlok := false;
  PGIBOX('Le libellé de la simulation doit être renseigné.', Ecran.caption);
  setfocuscontrol('LIBELLESIMUL');
  end;

  if ((typedepart = '') and ctlok)  then begin
  ctlok := false;
  PGIBOX('Le type de départ doit être renseigné.', Ecran.caption);
  setfocuscontrol('TYPEDEPART');
  end;

  if ((methcalcul = '') and ctlok) then begin
  ctlok := false;
  PGIBOX('La méthode de calcul des IDR doit être renseignée.', Ecran.caption);
  setfocuscontrol('METHCALCUL');
  end;

  if ((tauxactual = 0) and ctlok) then begin
  ctlok := false;
  PGIBOX('Le taux d''actualisation doit être renseigné.', Ecran.caption);
  setfocuscontrol('TAUXACTUAL');
  end;

  if ((codecalcul = '') and ctlok) then begin
  ctlok := false;
  PGIBOX('Le code calcul salaire doit être renseigné.', Ecran.caption);
  setfocuscontrol('CODECALCUL');
  end;

  // controle que le calcul des salaires pour cette date a été effectué avec le code calcul saisi
  Sql1 := 'Select PSY_CODECALSAL from SALAIREMOYEN where (PSY_CODECALSAL = "'+Codecalcul+'"'+
  ' and PSY_DATECALC = "'+usdatetime(datecalcul)+'")';
  if (not ExisteSql(Sql1) and ctlok) then
  begin
  ctlok := false;
  PGIBOX('Les salaires n''ont pas été calculés à la date du '+dtcalcul+' pour le code calcul '+Codecalcul+'.', Ecran.caption);
  setfocuscontrol('DATECALC');
  end;

  // si controles OK ==> lancement de la simulation
  if ctlok then begin
  St := dtsimul + ';' + codesimul +  ';' + libellesimul +';' + dtcalcul + ';' + typedepart + ';' + methcalcul +
  ';' + codecalcul + ';' + txactu + ';' + mode + ';' + codesimd ;

  // Appel écran de sélection des populations
  retour:= AGLLanceFiche('PAY','SIMULREGLEIDR_MUL','','',St);

  if retour = 'OK'  then
  begin
  TFvierge(Ecran).Retour := 'OK';
  TFVierge(Ecran).BFerme.Click;
  end
  else
  begin
  TFvierge(Ecran).retour := ' ';
  Setfocuscontrol('DATESIMUL');
  end;

  end;
  
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /    
Description .. : Recherche taux actualisation dans l'élément national saisi 
Suite ........ : dans les paramètres société
Mots clefs ... : 
*****************************************************************}
procedure TOF_CALCULSIMULPARAM.Recherchetauxactu ;
var
paramtaux : string;
QQ: Tquery;
dtsimul : Tdatetime;
Txactu : THNumEdit;

begin
paramtaux := vh_paie.pgidrelt;
 if paramtaux <> '' then
 begin

  dtsimul:=  StrToDate(getcontroltext('DATESIMUL'));
  // recherche du taux d'actualisation dans l'élément national
  QQ := opensql('select ##TOP 1## pel_montanteuro from eltnationaux where ##PEL_PREDEFINI## pel_codeelt="'+paramtaux+'" AND'+
  ' PEL_DATEVALIDITE <= "'+usdatetime(dtsimul)+'" Order by PEL_DATEVALIDITE desc', true);
  txactu := THNumEdit(GetControl('TAUXACTUAL'));
  if (txactu <> nil) and (not QQ.EOF) then txactu.Value := QQ.FindField('PEL_montanteuro').asFloat;
  ferme(QQ);
 end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/11/2006
Modifié le ... :   /  /    
Description .. : Codecalcul Exit
Mots clefs ... : 
*****************************************************************}
procedure TOF_CALCULSIMULPARAM.CodecalculExit(Sender: TObject);
var
Q: Tquery;
codecalcul : string;
begin
 codecalcul := getcontroltext('CODECALCUL');
 if codecalcul <> '' then
 begin
   Q:= opensql('SELECT PCY_LIBELLE from CALCULSALAIREMOY where PCY_CODECALSAL = "'+codecalcul+'"', true);
   if not Q.EOF then
    setcontroltext('LIBELLECALCUL', Q.findfield('PCY_LIBELLE').asstring)
   else
   begin
    setcontroltext('LIBELLECALCUL','');
    PGIBOX('Le code calcul salaire n''existe pas.', Ecran.caption);
    setfocuscontrol('CODECALCUL');
   end;
   ferme(Q);
 end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/11/2006
Modifié le ... :   /  /    
Description .. : Elipsis sur le champ CODE CALCUL
Mots clefs ... : 
*****************************************************************}
procedure TOF_CALCULSIMULPARAM.codecalculElipsisClick(Sender : TObject);
var St,storder ,codecalcul, dt : String;
Q : TQUERY;
begin
        St := 'SELECT DISTINCT PSY_DATECALC, PSY_CODECALSAL, PCY_LIBELLE FROM SALAIREMOYEN '+
        'LEFT JOIN CALCULSALAIREMOY ON PSY_CODECALSAL = PCY_CODECALSAL';

        Storder := 'SELECT DISTINCT PSY_DATECALC, PSY_CODECALSAL, PCY_LIBELLE FROM SALAIREMOYEN '+
        'LEFT JOIN CALCULSALAIREMOY ON PSY_CODECALSAL = PCY_CODECALSAL ORDER BY PSY_DATECALC DESC ';

        {$IFNDEF EAGLCLIENT}
        LookupList(THEdit(Sender),'Liste des codes calcul','SALAIREMOYEN LEFT JOIN CALCULSALAIREMOY'
        ,'PSY_CODECALSAL,PSY_DATECALC','PCY_LIBELLE','','',True,-1,storder);
        {$else}
        LookupList(THEdit(Sender),'Liste des codes calcul','SALAIREMOYEN LEFT JOIN CALCULSALAIREMOY'
        ,'PSY_CODECALSAL,PSY_DATECALC','PCY_LIBELLE','','PSY_DATECALC DESC',True,-1,st);
        {$ENDIF}

        dt := copy(getcontroltext('CODECALCUL'),5,10);
        codecalcul := copy(getcontroltext('CODECALCUL'),1,3);
        setcontroltext('CODECALCUL', codecalcul);
        setcontroltext('DATECALC', dt);
        if codecalcul <> '' then
        begin
         Q:= opensql('SELECT PCY_LIBELLE from CALCULSALAIREMOY where PCY_CODECALSAL = "'+codecalcul+'"', true);
         if not Q.EOF then setcontroltext('LIBELLECALCUL', Q.findfield('PCY_LIBELLE').asstring);
         ferme(Q);
        end
        else
        setcontroltext('LIBELLECALCUL', '');

end;

Initialization
  registerclasses ( [ TOF_CALCULSIMULPARAM ] ) ;
end.
