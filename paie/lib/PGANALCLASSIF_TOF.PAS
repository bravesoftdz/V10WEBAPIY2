{***********UNITE*************************************************
Auteur  ...... : 3GG
Créé le ...... : 22/02/2006
Modifié le ... : 22/02/2006
Description .. : Source TOF de la FICHE : PGANALCLASSIF ()
Mots clefs ... : TOF;PGANALCLASSIF
*****************************************************************}
Unit PGANALCLASSIF_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     QRS1,
     mul,
{$else}
     eMul, 
     UtileAGL,
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     paramdat, 
     UTOF ; 

Type
  TOF_PGANALCLASSIF = Class (TOF)
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

procedure TOF_PGANALCLASSIF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGANALCLASSIF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGANALCLASSIF.OnUpdate ;
var
  THDateDeb,THDateFin           : THEdit;
  DATEDEB,DATEFIN               : TDateTime;
  SQL,WHERE :STRING;
begin
  Inherited ;
  THDateDeb:= ThEdit(getcontrol('CRI_DATEDEBUT'));
  THDateFin:= ThEdit(getcontrol('CRI_DATEFIN'));
  DATEDEB:=strToDatetime(THDateDeb.text);
  DATEFIN:=StrToDatetime(ThDateFin.text);
  SQL:='SELECT ppu_etablissement,ppu_salarie,ppu_libelle,ppu_prenom,' +
  'ppu_indice,ppu_libelleemploi,ppu_qualification,ppu_convention,ppu_niveau,'+
  'ppu_coefficient,ppu_datedebut,ppu_datefin,PMI_LIBELLE,phb_mtrem' +
  'FROM paieencours' +
  'left join MINIMUMCONVENT on pmi_convention=ppu_convention AND PMI_NATURE="NIV" AND pmi_typenature = "VAL" AND pmi_code = ppu_NIVEAU' +
  'left join histobulletin on phb_etablissement=ppu_etablissement AND phb_salarie=ppu_salarie AND phb_datedebut=ppu_datedebut';
  WHERE := 'WHERE phb_rubrique="0002" AND phb_mtrem <> 0' +
  'AND ppu_datedebut >="'+Usdatetime(DATEDEB)+'" AND ppu_datefin <= "'+usdatetime(DATEFIN)+'"'+
  'ORDER BY ppu_etablissement,ppu_salarie,ppu_datedebut';
  TFQRS1(Ecran).WhereSQL:=SQL+WHERE;
end ;

procedure TOF_PGANALCLASSIF.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGANALCLASSIF.OnArgument (S : String ) ;
var
  THDateDeb,THDateFin           : THEdit;
begin
  Inherited ;
  THDateDeb:= ThEdit(getcontrol('CRI_DATEDEBUT'));
  THDateFin:= ThEdit(getcontrol('CRI_DATEFIN'));
end ;

procedure TOF_PGANALCLASSIF.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGANALCLASSIF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGANALCLASSIF.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_PGANALCLASSIF ] ) ; 
end.
