{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 31/05/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGHISTORETENUE ()
Mots clefs ... : TOF;PGHISTORETENUE
*****************************************************************}
Unit UTofPgHistoRetenue;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     UTob,forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF ;

Type
  TOF_PGHISTORETENUE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Implementation

procedure TOF_PGHISTORETENUE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTORETENUE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTORETENUE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTORETENUE.OnLoad ;
var T_Histo:Tob;
QHisto:TQuery;
Grille:THGrid;
begin
  Inherited ;
QHisto:=OpenSQL('SELECT PHR_DATEDEBUT,PHR_DATEFIN,CO_LIBELLE,PHR_MONTANT FROM HISTORETENUE'+
' LEFT JOIN COMMUN ON CO_CODE=PHR_TYPERGT AND CO_TYPE="PGT"',True);
T_Histo:=TOB.Create ('Historique des retenues', NIL, -1);
T_Histo.LoadDetailDB ('HISTORETENUE','', '', QHisto, FALSE) ;
Ferme(QHisto);
Grille:=THGrid(GetControl('GHISTO'));
T_Histo.PutGridDetail(Grille,false,false,
           'PHR_DATEDEBUT;PHR_DATEFIN;PHR_MONTANT;CO_LIBELLE',false);
T_Histo.Free;
end ;

procedure TOF_PGHISTORETENUE.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTORETENUE.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_PGHISTORETENUE ] ) ;
end.
