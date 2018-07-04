{***********UNITE*************************************************
Auteur  ...... : PAIE
Créé le ...... : 06/02/2002
Modifié le ... :   /  /    
Description .. : Gestion des groupes internes (unités de gestion)
Suite ........ : table GROUPEINTERNE
Suite ........ : fiche GROUPINTERNE
Suite ........ : 
Suite ........ : Le groupe interne est l'association d'une institution et d'une 
Suite ........ : catégorie de personnel. Permet la gestion des ruptures de 
Suite ........ : DUCS par groupe interne.
Suite ........ : 
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
{
PT1    : 16/10/2002 MF V585 Modification de la propriété Plus du champ
                            PGI_INSTITUTION - Corrige le PB : en mono dossier
                            la liste des institutions ne s'affichait pas.
}
unit UtomGroupeInterne;

interface

uses
    Classes, UTOM,PgOutils,HMsgBox, Controls ;

Type
      TOM_GROUPEINTERNE = Class (TOM)
        procedure OnArgument (stArguments : String ) ; override ;// PT1
        procedure OnUpdateRecord ; override ;
        procedure OnDeleteRecord; override;
      END;
implementation
// d PT1
procedure TOM_GROUPEINTERNE.OnArgument(stArguments: String);
begin

inherited ;
  SetControlProperty('PGI_INSTITUTION','Plus',' AND PIP_INSTITUTION not LIKE "Z%"');

end;
// f PT1
procedure TOM_GROUPEINTERNE.OnUpdateRecord;
var
Instit, Categ : string;
begin
  inherited;
Instit := GetField ('PGI_INSTITUTION');
if (Instit = '')  then
 begin
 LastError := 1;
 LastErrorMsg:='Vous devez renseigner le champ institution';
 SetFocusControl ('PGI_INSTITUTION');
 end;

Categ := GetField ('PGI_CATEGORIECRC');
if (Categ = '')  then
 begin
 LastError := 1;
 LastErrorMsg := 'Vous devez renseigner le champ catégorie ';
 SetFocusControl ('PGI_CATEGORIECRC');
 end;

end;

procedure TOM_GROUPEINTERNE.OnDeleteRecord;
var
NomChamp: array[1..1] of Hstring;
ValChamp: array[1..1] of variant;
ExisteCod : Boolean;
reponse : word;
begin
  inherited;

NomChamp[1]:='PDU_NUMERO'; ValChamp[1]:=GetField('PGI_GROUPE');
ExisteCod:=RechEnrAssocier('DUCSENTETE',NomChamp,Valchamp);
if ExisteCod=TRUE then
  begin
  LastError:=1;
  LastErrorMsg:='Attention! ce n° de groupe a déjà été utilisé pour la DUCS. Suppression impossible!';
  end;

NomChamp[1]:='PSA_DADSPROF'; ValChamp[1]:=GetField('PGI_CATEGORIECRC');
ExisteCod:=RechEnrAssocier('SALARIES',NomChamp,Valchamp);
if ExisteCod=TRUE then
  begin
  reponse := HShowMessage(';Confirmez-vous la suppression ?;le code catégorie est utilisé pour certains salariés; Q;YN;Y;N;;;','','');
  if (reponse <> mrYes) then
   LastError:=1;
   LastErrorMsg :='';
  end;

NomChamp[1]:='POG_RUPTGROUPE'; ValChamp[1]:='X';
ExisteCod:=RechEnrAssocier('ORGANISMEPAIE',NomChamp,Valchamp);
if ExisteCod=TRUE then
  begin
  reponse := HShowMessage('; Confirmez-vous la suppression ?;Certains organismes gèrent les ruptures sur groupe interne;Q;YN;Y;N;;;','','');
  if (reponse <> mrYes) then
   LastError:=1;
   LastErrorMsg :='';
  end;

end;
Initialization
registerclasses([TOM_GROUPEINTERNE]) ;
end.
 