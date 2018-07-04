{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 03/06/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGORDRERETENUE ()
Mots clefs ... : TOF;PGORDRERETENUE
*****************************************************************
---- JL 20/03/2006 modification clé annuaire ----
}
Unit UTofPgOrdreRetenue;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     sysutils,HCtrls,UTOF,UTob ;

Type
  TOF_PGORDRERETENUE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    Salarie:String;
  end ;

Implementation

procedure TOF_PGORDRERETENUE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGORDRERETENUE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGORDRERETENUE.OnUpdate ;
var Grille:THGrid;
    i:Integer;
    Valeur,NumOrdre:String;
begin
Inherited ;
Grille:=THGrid(GetControl('GRETENUE'));
If Grille<>Nil then
   begin
   i:=1;
   NumOrdre:=Grille.CellValues[0,1];
   While NumOrdre<>'' do
         begin
         Valeur:=Grille.CellValues[6,i];
         ExecuteSQL('UPDATE RETENUESALAIRE SET PRE_NIVEAURS='+Valeur+' WHERE PRE_SALARIE="'+Salarie+'" AND PRE_ORDRE='+NumOrdre+'');  //DB2
         i:=i+1;
         NumOrdre:=Grille.CellValues[0,i];
         end;
   end;
end ;

procedure TOF_PGORDRERETENUE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGORDRERETENUE.OnArgument (S : String ) ;
var T_Ordre:Tob;
    St,Rubrique:String;
    Q:TQuery;
    Grille:THGrid;
    i:Integer;
begin
Inherited ;
Salarie:=Trim(ReadTokenSt(S));
Rubrique:=S;
T_Ordre:= TOB.Create ('Les retenues', NIL,-1);
St:= 'SELECT ANN_NOMPER,PRE_DATEDEBUT,PRE_DATEFIN,PRE_MONTANTMENS,PRE_MONTANTTOT,PRE_NIVEAURS,PRE_ORDRE'+
     ' FROM RETENUESALAIRE LEFT JOIN ANNUAIRE ON PRE_BENEFRSGU=ANN_GUIDPER'+
     ' WHERE PRE_SALARIE="'+Salarie+'" AND PRE_RUBRIQUE="'+Rubrique+'" AND PRE_ACTIF="X"';
Q:=OpenSQL(St,TRUE);
T_Ordre.LoadDetailDB('RETENUESALAIRE','','',Q,False);
Ferme(Q);
Grille:=THGrid(GetControl('GRETENUE'));
T_Ordre.PutGridDetail (Grille, False, False, 'PRE_ORDRE;ANN_NOMPER;PRE_DATEDEBUT;PRE_DATEFIN;PRE_MONTANTMENS;PRE_MONTANTTOT;PRE_NIVEAURS',False) ;
T_Ordre.Free;
For i:=0 to 5 do
    begin
    Grille.ColEditables[i]:=False;
    end;
Grille.ColAligns[4]:=taRightJustify;
Grille.ColAligns[5]:=taRightJustify;
Grille.ColAligns[6]:=taCenter;
end ;

procedure TOF_PGORDRERETENUE.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_PGORDRERETENUE ] ) ;
end.
