{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGRECUPPREVISIONNEL ()
Mots clefs ... : TOF;PGRECUPPREVISIONNEL
*****************************************************************}
Unit UTofPgRecupPrevisionnel;

Interface

Uses StdCtrls,Classes,
     HCtrls,UTOF,Utob,PGHierarchie ;

Type
  TOF_PGRECUPPREVISIONNEL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    LeStage,LeMillesime,LaSession,MillesimeEC:String;
    end ;

Implementation

procedure TOF_PGRECUPPREVISIONNEL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGRECUPPREVISIONNEL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGRECUPPREVISIONNEL.OnUpdate ;
VAR Salarie,TypeResp,Champ:String;
    Tob1,Tob2,Tob3:Tob;
    MultiNiveau:Boolean;
begin
Inherited ;
If GetControlText('MULTI')='O' Then MultiNiveau:=True
Else MultiNiveau:=False;
Salarie:=GetControlText('SALARIE');
TypeResp:=GetControlText('TYPERESP');
Champ:=RenvoiChampResponsable(TypeResp);
EstResponsableHierarchique ( Salarie, TypeResp);
Tob1:=RecupServiceRespHierarchie (Salarie,TypeResp,MultiNiveau);
Tob2:=RecupSalariesHierarchie (Salarie,TypeResp,MultiNiveau);
Tob3:=RecupResponsablesHierarchie(Salarie,TypeResp,MultiNiveau);
Tob1.PutGridDetail(THGrid(GetControl('GRILLE1')),False,True,'PGS_CODESERVICE',False);
Tob2.PutGridDetail(THGrid(GetControl('GRILLE2')),False,True,'PSE_SALARIE',False);
Tob3.PutGridDetail(THGrid(GetControl('GRILLE3')),False,True,Champ,False);
Tob1.Free;
Tob2.Free;
Tob3.Free;
end ;

procedure TOF_PGRECUPPREVISIONNEL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGRECUPPREVISIONNEL.OnArgument (S : String ) ;
begin
Inherited ;
LeStage:=ReadTokenPipe(S,';');
LeMillesime:=ReadTokenPipe(S,';');
LaSession:=ReadTokenPipe(S,';');
MillesimeEC:=ReadTokenPipe(S,';');
end ;

procedure TOF_PGRECUPPREVISIONNEL.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_PGRECUPPREVISIONNEL ] ) ;
end.
