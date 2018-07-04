{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/03/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SALMOYENNONCAL ()
Mots clefs ... : TOF;SALMOYENNONCAL
*****************************************************************}
Unit SALMOYENNONCAL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
     {$IFNDEF EAGLCLIENT}
     db, 
     dbtables,
     HDB,
     mul,
     FE_Main,
     {$else}
     eMul,
     MaineAgl,
     {$ENDIF}
     uTob,
     sysutils,
     AglInit ,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97,
     HQRY,
     P5Def,
     UTOF ; 

Type
  TOF_SALMOYENNONCAL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    datecal : Tdatetime;
    lancecalcul : boolean;
    procedure Calculsal(Sender : Tobject);
  end ;

Implementation

procedure TOF_SALMOYENNONCAL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_SALMOYENNONCAL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_SALMOYENNONCAL.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 10/07/2006
Modifié le ... :   /  /    
Description .. : Chargement
Mots clefs ... : 
*****************************************************************}
procedure TOF_SALMOYENNONCAL.OnLoad ;
var
datesaisi, stwhere : string;
begin
  Inherited ;
  datesaisi := getcontrolText('DATECALCUL');
  Datecal := StrToDate(Datesaisi);
  if Lancecalcul then exit;
  Stwhere := '(PSA_DATESORTIE >="'+usdatetime(datecal)+'" or'+
  ' PSA_DATESORTIE <="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) and (PSA_DATEENTREE <= "'+
  usdatetime(datecal)+'") and (PSA_SALARIE not in (select PSY_SALARIE from SALAIREMOYEN where'+
   ' PSY_DATECALC ="'+usdatetime(datecal)+'"))';
   Setcontroltext('XX_WHERE', stwhere);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 10/07/2006
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_SALMOYENNONCAL.OnArgument (S : String ) ;
VAR
Anneeprec, num : integer;
AnneeA, MoisM, jour : word;
jourj, zdate : Tdatetime;
Btn : TToolBarButton97;
begin
  Inherited ;
  Jourj := Date;
  decodedate(Jourj,AnneeA, MoisM, Jour);
  Anneeprec := AnneeA - 1;
  zdate := Encodedate(anneeprec, moisM, Jour);
  datecal := finannee(Zdate);
  SetControlText ('DATECALCUL', DateTimeToStr(datecal));

  for Num := 1 to 4 do
  begin
  VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));


  Btn := TToolBarButton97(GetControl('BOUVRIR'));
  if btn <> nil then Btn.Onclick := Calculsal;

end ;

procedure TOF_SALMOYENNONCAL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_SALMOYENNONCAL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_SALMOYENNONCAL.OnCancel () ;
begin
  Inherited ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 10/07/2006
Modifié le ... :   /  /    
Description .. : Calcul des salaires moyens
Mots clefs ... : 
*****************************************************************}
procedure TOF_SALMOYENNONCAL.Calculsal (sender : Tobject);
var
Salarie, st , lewhere, param: string;
i: integer;


begin

      if (TFmul(Ecran).Fliste.NbSelected = 0) and (not TFmul(Ecran).Fliste.AllSelected) then
      begin
      PGIBOX('Aucun salarié sélectionné', Ecran.Caption);
      exit;
      end;

  { Gestion de la sélection des salariés}

    if PgiAsk('Calcul des salaires moyens pour les salariés sélectionnés. Voulez-vous poursuivre ?', Ecran.caption) = mrYes then
    begin  // D2

     St := '';
    if (TFmul(Ecran).Fliste.AllSelected = false) then

     // lecture de chaque salarié sélectionné
     begin
      { Composition du clause WHERE pour limiter le mul à ces salariés }
      for i := 0 to TFmul(Ecran).Fliste.NbSelected - 1 do
      begin
{$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
        TFMul(Ecran).Fliste.GotoLeBOOKMARK(i);
        Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
        St := St + ' PSA_SALARIE="' + Salarie + '" OR';
      end;
      TFMul(Ecran).FListe.ClearSelected;
     end;

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
  param := 'N' + ';' + lewhere;
  AglLanceFiche('PAY', 'CALCULSALMOYEN', '', '', param);
  TheMulQ := nil;

  TFMul(Ecran).BCherche.Click;
   end; // E2

  end;

Initialization
  registerclasses ( [ TOF_SALMOYENNONCAL ] ) ; 
end.
