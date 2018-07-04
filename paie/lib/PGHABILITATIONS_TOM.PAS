{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 09/01/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PGHABILITATIONS (PGHABILITATIONS)
Mots clefs ... : TOM;PGHABILITATIONS
*****************************************************************}
Unit PGHABILITATIONS_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList, 
{$else}
     eFiche, 
     eFichList, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     pgoutils,
     UTOM,
     UTob ;

Type
  TOM_PGHABILITATIONS = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;

    private
     CEG, STD, DOS : boolean;
     procedure DoChanged(sender : TObject);
     procedure Modiftablettepopul(sender : TObject);
    end ;

Implementation
    uses PgOutils2; // appel à la fonction PgrendNodossier pour affectation n° de dossier.


{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : On update record
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGHABILITATIONS.OnUpdateRecord ;
var
sql, pop : string;
QQ : TQUERY;
begin
  Inherited ;


   if (GetControlText('PHL_GRPUTI')='UTI') and (GetControlText('PHL_USER') = '') then
   begin
     LastError := 1;
     PGIBOX('Un utilisateur doit être renseigné.','Saisie des habilitations');
     SetFocusControl('PHL_USER');
   end
   else if (GetControlText('PHL_GRPUTI')='GRP') and (GetControlText('PHL_USERGRP') = '') then
   begin
     LastError := 1;
     PGIBOX('Un groupe d''utilisateur doit être renseigné.','Saisie des habilitations');
     SetFocusControl('PHL_USERGRP');
   end;
   {$IFNDEF EAGLCLIENT}
  pop := THDBValcomboBox(getcontrol('PHL_POPULATION')).value;
   {$ELSE}
  pop := THValcomboBox(getcontrol('PHL_POPULATION')).value;
   {$ENDIF}
  if pop = '' then
  begin
     LastError := 1;
     PGIBOX('La population doit être renseignée.','Saisie des habilitations');
     SetFocusControl('PHL_POPULATION');
  end
  else
  begin
     sql := 'Select PPC_CODEPOP from ORDREPOPULATION where PPC_POPULATION = "'+pop+'"';
     if not existesql(sql) then
     begin
     LastError := 1;
     PGIBOX('Cette population n''existe pas.','Saisie des habilitations');
     SetFocusControl('PHL_POPULATION');
     end;
  end;
  sql := 'Select PPC_CODEPOP from ORDREPOPULATION where PPC_POPULATION = "'+pop+'"';
  QQ := opensql(sql, true);
  if not QQ.EOF then setfield('PHL_CODEPOP', QQ.findfield('PPC_CODEPOP').asstring);
  ferme(QQ);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : Chargement de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGHABILITATIONS.OnLoadRecord ;
 var str_infosys,str_datecreate,str_dateupdate : variant;
begin
  Inherited ;
  //Affiche des infos relatifs aux champs systèmes.date de création, utilisateur....
  try
    str_datecreate := FormatDateTime('dd/mm/yyyy hh:mm',GetField('PHL_DATECREATION'))
  except
    str_datecreate := '';
  end;

  try
    str_dateupdate := FormatDateTime('dd/mm/yyyy hh:mm',GetField('PHL_DATEMODIF'))
  except
    str_dateupdate := '';
  end;

  str_infosys := 'Créé le '+str_datecreate+' par '+GetField('PHL_CREATEUR')+' / '
                +'Modifié le '+str_dateupdate+' par '+GetField('PHL_UTILISATEUR');

  SetControlText('TPHL_INFOSYS',str_infosys);

  // init en mode insertion
  if ds.State = DsInsert then
  begin
    SetField('PHL_GRPUTI','UTI');
    SetFocusControl('PHL_PREDEFINI');
    SetControlEnabled('TPHL_USERGRP',FALSE);
    SetControlEnabled('PHL_USERGRP',FALSE);
  end
  else
    SetControlEnabled('PHL_GRPUTI',FALSE);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : Modification d'un champ
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGHABILITATIONS.OnChangeField ( F: TField ) ;
var
pred : string;

begin
  Inherited ;

  if (F.FieldName = 'PHL_PREDEFINI') and (ds.State = DsInsert) then
  begin

    Pred := GetField('PHL_PREDEFINI');

    if Pred = '' then exit;

    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer d''habilitation prédéfinie CEGID.', 'Saisie des habilitations');
      Pred := 'DOS';
      SetControlProperty('PHL_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer d''habilitation prédéfinie Standard.', 'Saisie des habilitations');
      Pred := 'DOS';
      SetControlProperty('PHL_PREDEFINI', 'Value', Pred);
    end;
     // affectation du n° de dossier en fonction du prédéfini
    if GetControlText('PHL_PREDEFINI')='DOS' then
       SetField('PHL_NODOSSIER', PgrendNodossier())
    else
        SetField('PHL_NODOSSIER', '000000');

  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGHABILITATIONS.OnArgument ( S: String ) ;
var
predefini , GrpUtiCmbbx : THValComboBox;

begin
  Inherited ;

  AccesPredefini('TOUS', CEG, STD, DOS);

  GrpUtiCmbbx := THValComboBox(GetControl('PHL_GRPUTI',TRUE));
  if GrpUtiCmbbx <> nil then
    GrpUtiCmbbx.OnChange := DoChanged;

   
  predefini := THValcomboBox(getcontrol('PHL_PREDEFINI', TRUE));
  if Predefini<> nil then predefini.Onchange := Modiftablettepopul;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : Initialise les champs utilisateur/groupe utilisateur
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGHABILITATIONS.DoChanged(sender: TObject);
begin
  // Active ou Désactive les combobox en fonction du critère GROUPE ou UTILISATEUR.
  if (sender as THDBValComboBox).name = 'PHL_GRPUTI' then
  begin
    if GetControlText('PHL_GRPUTI') = 'GRP' then
    begin
      SetControlEnabled('TPHL_USER',FALSE);
      SetControlEnabled('PHL_USER',FALSE);
      if GetControlText('PHL_USER') <> '' then
        SetControlText('PHL_USER','');

      SetControlEnabled('TPHL_USERGRP',TRUE);
      SetControlEnabled('PHL_USERGRP',TRUE);
    end
    else if GetControlText('PHL_GRPUTI') = 'UTI' then
    begin
      SetControlEnabled('TPHL_USER',TRUE);
      SetControlEnabled('PHL_USER',TRUE);

      SetControlEnabled('TPHL_USERGRP',FALSE);
      SetControlEnabled('PHL_USERGRP',FALSE);
      if GetControlText('PHL_USERGRP') <> '' then
        SetControlText('PHL_USERGRP','');
    end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : Initialise la tablette des populations selon le prédéfini saisi
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGHABILITATIONS.Modiftablettepopul(sender: TObject);
var
predefini,  pred, sql1 , nodossier, st: string;

begin
  Setcontroltext('PHL_POPULATION', '');
  predefini := THValcomboBox(getcontrol('PHL_PREDEFINI')).value;
  nodossier := PgrendNodossier();


  if predefini = 'STD' then      // si prédefini saisi = 'Standard"
  begin
     // recherche des populations de type STD et s'il n'en existe pas : recherche pop CEGID
     pred := 'STD';
     sql1 := 'select PPC_POPULATION,PPO_CODEPOP from ORDREPOPULATION, CODEPOPULATION WHERE PPC_CODEPOP = PPO_CODEPOP'+
     ' and PPO_PREDEFINI = "STD" and  PPO_TYPEPOP like "%HAB%"';
     if not existesql(sql1) then  pred := 'CEG';

  end
  else
  begin
        // si prédéfini = Dossier
    if predefini = 'DOS' then
    begin
      // recherche des populations de type Dossier ou STD  ou CEGID
      pred := 'DOS';
      sql1 := 'select PPC_POPULATION,PPO_CODEPOP from ORDREPOPULATION,CODEPOPULATION WHERE PPC_CODEPOP = PPO_CODEPOP'+
      ' and PPO_PREDEFINI = "DOS" and PPO_NODOSSIER = "'+nodossier+'" and PPO_TYPEPOP like "%HAB%"';
      if not existesql(sql1) then
      begin
        // recherche si population existe en standard
        pred := 'STD';
        sql1 := 'select PPC_POPULATION,PPO_CODEPOP from ORDREPOPULATION, CODEPOPULATION WHERE PPC_CODEPOP = PPO_CODEPOP'+
        ' and PPO_PREDEFINI = "STD" and PPO_TYPEPOP like "%HAB%"';
        if not existesql(sql1) then   pred := 'CEG';
      end;
    end;
  end;

    st := ' and PPC_PREDEFINI = "'+pred+'"';
    setcontrolproperty('PHL_POPULATION', 'plus', st);

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 02/04/2007
Modifié le ... :   /  /    
Description .. : Création
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGHABILITATIONS.OnNewRecord ;
begin
  Inherited ;
  if (STD = TRUE) then
    SetField('PHL_PREDEFINI', 'STD')
  else
    SetField('PHL_PREDEFINI', 'DOS');
  Modiftablettepopul(nil);
end ;

procedure TOM_PGHABILITATIONS.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_PGHABILITATIONS.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_PGHABILITATIONS.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_PGHABILITATIONS.OnClose ;
begin
  Inherited ;
 
end ;

procedure TOM_PGHABILITATIONS.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [TOM_PGHABILITATIONS] ) ;
end.
