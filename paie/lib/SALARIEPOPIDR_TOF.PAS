{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 13/06/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SALARIEPOPIDR ()
                 Salariés affectés à une population
Mots clefs ... : TOF;SALARIEPOPIDR
*****************************************************************}
Unit SALARIEPOPIDR_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     HDB,
{$else}
     eMul,
     uTob,
{$ENDIF}
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     UTOF ;

Type
  TOF_SALARIEPOPIDR = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private

    Population , codepop: string;
    end ;

Implementation
uses P5DEF,
     ENTPAIE,
     pgpopuloutils;

procedure TOF_SALARIEPOPIDR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_SALARIEPOPIDR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_SALARIEPOPIDR.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 15/06/2006
Modifié le ... :   /  /
Description .. : Chargement des enregistrements
Mots clefs ... :
*****************************************************************}
procedure TOF_SALARIEPOPIDR.OnLoad ;
VAR
DateArret, finperdate : TdateTime;
StDateArret, sql1, st1,st2, finper : string;
histo, paieencours : boolean;
Q1 : TQuery;

begin
  Inherited ;
  paieencours := false;
  DateArret:=StrtoDate(GetControlText('DATEARRET'));
// recherche si la date arrêt est > date fin dernière paie
  Q1 := opensql('SELECT PEX_FINPERIODE FROM EXERSOCIAL WHERE PEX_ACTIF = "X" ORDER BY PEX_ANNEEREFER DESC', true);
  if not Q1.EOF then
  begin
    finper := Q1.findfield('PEX_FINPERIODE').asstring;
    finperdate := StrtoDate(finper);
    if Datearret > finperdate then paieencours := false else paieencours := true;
  end;
  ferme(Q1);



// Constitution de la clause WHERE
  setcontroltext('XX_WHERE', '');

 // recherche de la clause Where pour cette population

 histo := false;
  if (paieencours  and vh_paie.pghistorisation) then  histo := true;

  PGrechwhere(Codepop, population, histo, st1,st2);

  if (histo = false) or (st2 = '') then
  begin
  TFMUL(Ecran).setDbliste('PGSALARIES');

// recherche des donnes SALARIES avec la clause where SALARIES

  sql1 := st1;
  end
  else
   begin
   TFMUL(Ecran).setDBliste('PGSALARIESPOP');
   if st1 <> '' then
    sql1 := ''+st1+' and  a.Phs_dateapplic = (select max(phs_dateapplic) from HISTOSALARIE b'+
    ' Where b.phs_dateapplic<= "'+UsdateTime(datearret)+'" and a.phs_salarie= b.phs_salarie) and '+st2+''
   else
    sql1 := 'a.Phs_dateapplic = (select max(phs_dateapplic) from HISTOSALARIE b'+
    ' Where b.phs_dateapplic<= "'+UsdateTime(datearret)+'" and a.phs_salarie= b.phs_salarie) and '+st2+'';
  end;



   // si salarié sorti sélectionné :
  if  TCheckBox(GetControl('CKSORTIE')) <> nil then
  Begin
   if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then
     Begin
     StDateArret:=' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
     StDateArret:=StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
     Sql1 := Sql1 + StDateArret;
     End
   End;

    SetControlText('XX_WHERE',Sql1);

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 13/06/2006
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_SALARIEPOPIDR.OnArgument (S : String ) ;
var
Libpop : string;     // libellé de la population
num : integer;


begin
  Inherited ;
codepop := Readtokenst(S);
population := ReadTokenSt(S);
libpop := ReadTokenSt(S);

TFMul(Ecran).Caption := 'Salariés affectés à la population ' + population + ' ' + libpop;

UpdateCaption(TFMul(Ecran));


// Recherche du libellé des zones TRAVAIL et STAT
  for Num := 1 to 4 do
  begin
  VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));

 

end ;


procedure TOF_SALARIEPOPIDR.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_SALARIEPOPIDR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_SALARIEPOPIDR.OnCancel () ;
begin
  Inherited ;
end ;



Initialization
  registerclasses ( [ TOF_SALARIEPOPIDR ] ) ;
end.
