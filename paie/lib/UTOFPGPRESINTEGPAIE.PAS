{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 30/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PRESINTEGPAIE () Présence intégrée paie
Mots clefs ... : TOF;INTEGPREPAIE
*****************************************************************}
Unit UTOFPGPRESINTEGPAIE ;

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
//     uTob,
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
     Utob,
     Hqry,
     HStatus,
     ed_tools,
     AglInit;


Type
  TOF_PGPRESINTEGPAIE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
 private

    DateDeb, DateFin: THedit;
    procedure MiseEnFormeXXwhere;
    procedure Datedebexit(sender : Tobject);
    procedure Datefinexit(sender : Tobject);

 end;

Implementation
uses PgOutils2;

procedure TOF_PGPRESINTEGPAIE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGPRESINTEGPAIE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGPRESINTEGPAIE.OnUpdate ;
begin
  Inherited ;
end ;


procedure TOF_PGPRESINTEGPAIE.OnLoad ;
begin
  Inherited ;
  MiseenFormeXXWhere;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 30/10/2007
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGPRESINTEGPAIE.OnArgument (S : String ) ;
var
Btn : TToolBarButton97;
DebPer, FinPer, ExerPerEncours : string;
OkOk: Boolean;
num : integer;
typeaffichage: THValComboBox;
begin
  Inherited ;


  for Num := 1 to 4 do
  begin
  VisibiliteChampSalarie(IntToStr(Num), GetControl('PYP_TRAVAILN' + IntToStr(Num)), GetControl('TPYP_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PYP_CODESTAT'), GetControl('TPYP_CODESTAT'));


  DateDeb := ThEdit(getcontrol('PERIODEDEBUT'));
  if DateDeb <> nil then DateDeb.OnExit := DateDebExit;
  DateFin := ThEdit(getcontrol('PERIODEFIN'));
  if DateFin <> nil then DateFin.OnExit := DateFinExit;

  OkOk := RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer);
  if OkOk then
  begin
    if DateDeb <> nil then DateDeb.text := DebPer;
    if DateFin <> nil then DateFin.text := FinPer;
  end;

end ;

procedure TOF_PGPRESINTEGPAIE.OnClose ;
begin
  Inherited ;

end ;

procedure TOF_PGPRESINTEGPAIE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGPRESINTEGPAIE.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 22/10/2007
Modifié le ... :   /  /    
Description .. : Sortie de date début exercice
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGPRESINTEGPAIE.DateDebExit(Sender: TObject);
begin
  if not IsValidDate(DateDeb.Text) then
  begin
    PGIBox('La date de début n''est pas valide', Ecran.Caption);
    DateDeb.SetFocus;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 22/10/2007
Modifié le ... :   /  /    
Description .. : Sortie de fin exercice
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGPRESINTEGPAIE.DateFinExit(Sender: TObject);
var
  Date1, Date2: TDateTime;
begin
  if not IsValidDate(DateFin.Text) then
  begin
    PGIBox('La date de fin n''est pas valide', Ecran.Caption);
    DateFin.SetFocus;
    exit;
  end;
  Date2 := StrToDate(DateFin.Text);
  Date1 := StrToDate(DateDeb.Text);
  if Date1 > Date2 then
  begin
    PGIBox('La date de début est supérieure à la date de fin', Ecran.Caption);
    DateDeb.SetFocus;
    DateFin.Text := '';
    exit;
  end;
 // if (FINDEMOIS(Date1) <> FINDEMOIS(Date2)) then
 //  begin
 //   PGIBox('Les dates de début et de fin doivent être comprise dans le même mois', Ecran.Caption);
 //   DateFin.SetFocus;
 //   exit;
//  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 30/10/2007
Modifié le ... :   /  /    
Description .. : Mise en forme XXWHERE
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGPRESINTEGPAIE.MiseenFormeXXWHERE;
var
Date1, Date2 : TDateTime;
Begin

    Date2 := StrToDate(DateFin.Text);
    Date1 := StrToDate(DateDeb.Text);
    SetControlText('XX_WHERE','PYP_PGINDICATPRES = "INP" AND PYP_DATEDEBUTBUL >="'+UsDateTime(Date1)+'"'+
    ' AND PYP_DATEFINBUL <= "'+Usdatetime(Date2)+'"');


end;



Initialization
  registerclasses ( [ TOF_PGPRESINTEGPAIE ] ) ;
end.
