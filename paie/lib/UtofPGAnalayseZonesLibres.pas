{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGANALYSEZONELIBRE ()
Mots clefs ... : TOF;PGANALYSEZONELIBRE
*****************************************************************
PT1     25/09/2007 FC V_80 FQ 14807 Analyse des éléments dynamiques et des éléments nationaux
}
Unit UtofPGAnalayseZonesLibres ;

Interface

Uses StdCtrls, 
     Controls, 
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
     HTB97,
     HMsgBox,
     PgOutilsHistorique,
     UTOF,
     ParamSoc;  //PT1

Type
  TOF_PGANALYSEZONELIBRE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    sDateValidite,sMode : String; //PT1
    procedure BMajClick(Sender : TObject);
    procedure OnClickSalarieSortie(Sender: TObject); //PT1
    procedure ActiveWhere(Sender: TObject); //PT1
    procedure ExitEdit(Sender: TObject); //PT1
  end ;

Implementation

uses
  PgOutils2;

procedure TOF_PGANALYSEZONELIBRE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSEZONELIBRE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSEZONELIBRE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSEZONELIBRE.OnLoad ;
begin
  Inherited ;
  //DEB PT1
  if (GetControlText('DATEVALIDITE') <> sDateValidite) or (GetControlText('CKLIBELLE') <> sMode) then
  begin
    if (GetControlText('CKLIBELLE') = '-') then
      PGMajDonneesAfficheZL(False,StrToDateTime(GetControlText('DATEVALIDITE')),True)
    else
      PGMajDonneesAfficheZL(False,StrToDateTime(GetControlText('DATEVALIDITE')),False);
    sDateValidite := GetControlText('DATEVALIDITE');
    sMode := GetControlText('CKLIBELLE');
  end;
  ActiveWhere(nil);
  //FIN PT1
end ;

procedure TOF_PGANALYSEZONELIBRE.OnArgument (S : String ) ;
var
  iTable, iChamp: integer;
  St, Prefixe, TheEtab, LesEtab: string;
  Q: TQuery;
  TOB_DesChamps: TOB;
  T1: TOB;
  zz, i: Integer;
  LeType, LaValeur, LePlus, LeChamp: string;
  Test : String;
  Bt : TToolBarButton97;
  Check : TCheckBox;  //PT1
  Edit : THEdit;     //PT1
begin
   Inherited;
   sDateValidite := Copy(DateTimeToStr(Now),1,10);
   sMode := '-';
   Prefixe := 'PTZ';
   iTable := PrefixeToNum(Prefixe);
   ChargeDeChamps(iTable, Prefixe);
   Q := OpenSQL('SELECT * FROM PGPARAMAFFICHEZL',True);
   for iChamp := 1 to high(V_PGI.DeChamps[iTable]) do
   begin
    For i := 1 to 30 do
    begin
      if V_PGI.DEChamps[iTable, iChamp].Nom = 'PTZ_PGVALZL'+IntToStr(i) then
      begin
        LaValeur := Q.FindField('PAZ_CHAMPDISPO'+IntToStr(i)).AsString;
        LeChamp := Copy(LaValeur,4,Length(LaValeur));
        LeType := Copy(LaValeur,1,3);
        If LeType = 'NAT' then
        begin
          V_PGI.DEChamps[iTable, iChamp].Libelle := RechDom('PGELEMENTNAT',LeChamp,False);
          V_PGI.DEChamps[iTable, iChamp].Control := 'LDC';
        end
        else If LeType = 'ZLS' then
        begin
          V_PGI.DEChamps[iTable, iChamp].Libelle := RechDom('PGZONEHISTOSAL',LeChamp,False);
          V_PGI.DEChamps[iTable, iChamp].Control := 'LDC';
        end
        else
        begin
          V_PGI.DEChamps[iTable, iChamp].Control := '';
          V_PGI.DEChamps[iTable, iChamp].Libelle := 'Champ libre non paramétré';
        end;
      end;
    end;
   end;
   Ferme(Q);
   Bt := TToolBarButton97(GetControl('BMAJDONNEES'));
   If Bt <> Nil then Bt.OnClick := BMajClick;
   //DEB PT1
   Check := TCheckBox(GetControl('CKSORTIE'));
   if (Check <> nil) then
     Check.OnClick := OnClickSalarieSortie;
   Edit := THEdit(GetControl('PSA_SALARIE'));
   if (Edit <> nil) then
     Edit.OnExit := ExitEdit;
   //FIN PT1
end ;

procedure TOF_PGANALYSEZONELIBRE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSEZONELIBRE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSEZONELIBRE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSEZONELIBRE.BMajClick(Sender : TObject);
begin
  //PT1
  if (GetControlText('CKLIBELLE') = '-') then
    PGMajDonneesAfficheZL(False,StrToDateTime(GetControlText('DATEVALIDITE')),True)
  else
    PGMajDonneesAfficheZL(False,StrToDateTime(GetControlText('DATEVALIDITE')),False);
  sMode := GetControlText('CKLIBELLE');
end;

//DEB PT1
procedure TOF_PGANALYSEZONELIBRE.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (GetParamSocSecur('SO_PGTYPENUMSAL','') = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGANALYSEZONELIBRE.OnClickSalarieSortie(Sender: TObject);
begin
  SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
  SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

procedure TOF_PGANALYSEZONELIBRE.ActiveWhere(Sender: TObject);
var
  StWhere : String;
  DateArret : TDateTime;
  StDateArret : String;
  i : integer;
  St : String;
begin
  SetControlText('XX_WHERE','');

  stWhere := '';
  // Critère matricule salarié
  if (GetControlText('PSA_SALARIE') <> '') then
    stWhere := stWhere + ' AND PSA_SALARIE = "' + GetControlText('PSA_SALARIE') + '"';

  // Critère nom salarié
  if (GetControlText('PSA_LIBELLE') <> '') then
    stWhere := stWhere + ' AND PSA_LIBELLE LIKE "' + GetControlText('PSA_LIBELLE') + '%"';

  // Critère établissement
  if (GetControlText('PSA_ETABLISSEMENT') <> '') and (GetControlText('PSA_ETABLISSEMENT') <> '<<Tous>>') then
    stWhere := stWhere + ' AND PSA_ETABLISSEMENT = "' + GetControlText('PSA_ETABLISSEMENT') + '"';

  // Gérer la case à cocher exclure
  if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET'))) then
  Begin
    DateArret := StrtoDate(GetControlText('DATEARRET'));
    StDateArret := ' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
    StDateArret := StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
    stWhere := stWhere + StDateArret;
  End;

  // Zone de travail 1 à 4
  for i := 1 to 4 do
  begin
    if (GetControlText('PSA_TRAVAILN' + IntToStr(i)) <> '') and (GetControlText('PSA_TRAVAILN' + IntToStr(i)) <> '<<Tous>>') then
      stWhere := stWhere + ' AND PSA_TRAVAILN' + IntToStr(i) + ' = "' + GetControlText('PSA_TRAVAILN' + IntToStr(i)) + '"';
  end;

  // Code statistique
  if (GetControlText('PSA_CODESTAT') <> '') and (GetControlText('PSA_CODESTAT') <> '<<Tous>>') then
    stWhere := stWhere + ' AND PSA_CODESTAT = "' + GetControlText('PSA_CODESTAT') + '"';

  St := SQLConf('SALARIES');
  if St <> '' then
    St := ' AND ' + St;
  stWhere := stWhere + St;

  SetControlText('XX_WHERE',stWhere);
end;
//FIN PT1

Initialization
  registerclasses ( [ TOF_PGANALYSEZONELIBRE ] ) ; 
end.

