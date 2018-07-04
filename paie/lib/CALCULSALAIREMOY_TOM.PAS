{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 09/05/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CALCULSALAIREMOY (CALCULSALAIREMOY)
Mots clefs ... : TOM;CALCULSALAIREMOY
*****************************************************************}
Unit CALCULSALAIREMOY_TOM ;

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
UTOB,
     eFichList,
{$ENDIF}
     forms, 
     sysutils,
     PgOutils2, 
     ComCtrls,
     HCtrls, 
     HMsgBox,
     PgOutils,
     UTOM;


Type
  TOM_CALCULSALAIREMOY = Class (TOM)
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
    LectureSeule, CEG, STD, DOS : boolean;
    end ;

Implementation

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 05/06/2006
Modifié le ... :   /  /    
Description .. : Nouvelle saisie
Mots clefs ... : 
*****************************************************************}
procedure TOM_CALCULSALAIREMOY.OnNewRecord ;
begin
  Inherited ;
  SetControlEnabled('methodecal',False);
if (CEG = TRUE) then
    SetField('PCY_PREDEFINI', 'CEG')
  else
    SetField('PCY_PREDEFINI', 'DOS');
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 05/06/2006
Modifié le ... :   /  /    
Description .. : Suppression du code calcul
Mots clefs ... : 
*****************************************************************}
procedure TOM_CALCULSALAIREMOY.OnDeleteRecord ;
var
sql1, sql : string;
begin
  Inherited ;
  // contrôle si ce code calcul a déjà été utilisé dans une simulation
  Sql := 'Select PSL_CODECALSAL from SIMULATION where PSL_CODECALSAL = "'+Getfield('PCY_CODECALSAL')+'"';
  if ExisteSql(Sql) then
  begin
  PGIBOX('Suppression interdite : Ce code calcul a déjà été utilisé .#13#10'+
  'pour une simulation.', Ecran.Caption);
  Lasterror := 1;
  exit;
  end;

  // Contrôle si ce code calcul a déjà été utilisé pour calculer le salaire moyen
  Sql1 := 'Select PSY_CODECALSAL from SALAIREMOYEN where PSY_CODECALSAL = "'+Getfield('PCY_CODECALSAL')+'"';
  if ExisteSql(Sql1) then
  begin
  PGIBOX('Suppression interdite : Ce code calcul a déjà été utilisé .#13#10'+
  'pour le calcul des salaires moyens.', Ecran.Caption);
  Lasterror := 1;
  exit;
  end;

  ExecuteSql('Delete from METHCALCULSALMOY where ##PSM_PREDEFINI## PSM_CODECALSAL ="'+Getfield('PCY_CODECALSAL')+'"');
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 05/06/2006
Modifié le ... :   /  /    
Description .. : Mise à jour
Mots clefs ... : 
*****************************************************************}
procedure TOM_CALCULSALAIREMOY.OnUpdateRecord ;
var
codecalsal, predefini : string;
i : integer;
codeok : boolean;

begin
  Inherited ;
  // contrôle du code calcul
codecalsal := Getfield('PCY_CODECALSAL');
predefini := Getfield('PCY_PREDEFINI');
if length(Codecalsal) <> 3 then begin
  PGIBOX('Le code calcul doit comporter 3 caractères.', Ecran.Caption);
  setfocuscontrol('PCY_CODECALSAL');
  Lasterror := 1;
  exit;
  end;

  for i := 1 to 3 do
  begin

  if (codecalsal[i] in ['0'..'9'])  or (codecalsal[i] in ['A'..'Z']) or (codecalsal[i] in ['a'..'z']) then
  codeOK := true
  else
  codeOK:= false;
  if not codeOK then break;
  end;

  if not codeOK then
  begin
  PGIBOX('Le code calcul doit comporter uniquement des chiffres ou des lettres.', Ecran.caption);
  setfocuscontrol('PCY_CODECALSAL');
  Lasterror := 1;
  exit;
  end;

if ((predefini = 'STD') and (Codecalsal[3] <> '1') and (Codecalsal[3] <> '3'))
  then begin
  PGIBOX('Le code calcul doit se terminer par 1 ou 3.', Ecran.Caption);
  setfocuscontrol('PCY_CODECALSAL');
  Lasterror := 1;
  exit;
  end;

if ((predefini = 'CEG') and (Codecalsal[3] <> '0') and (Codecalsal[3] <> '2') and (codecalsal[3] <> '4') and
   (codecalsal[3] <> '6') and (codecalsal[3] <> '8'))
  then begin
  PGIBOX('Le code calcul doit se terminer par 0,2,4,6 ou 8.', Ecran.Caption);
  setfocuscontrol('PCY_CODECALSAL');
  Lasterror := 1;
  exit;
  end;

If ((predefini = 'DOS') and (Codecalsal[3] <> '5') and (Codecalsal[3] <> '7') and (Codecalsal[3] <> '9'))
  then begin
  PGIBOX('Le code calcul doit se terminer par 5,7 ou 9.', Ecran.Caption);
  setfocuscontrol('PCY_CODECALSAL');
  Lasterror := 1;
  exit;
  end;

if (getfield('PCY_CHAMPAPPLIC') = '')
  then begin
  PGIBOX('Le champ application doit être renseigné.', Ecran.Caption);
  setfocuscontrol('PCY_CHAMPAPPLIC');
  Lasterror := 1;
  exit;
  end;
  if predefini = 'DOS' then setfield('PCY_NODOSSIER', PgrendNodossier())  else
  setfield('PCY_NODOSSIER', '000000');

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 05/06/2006
Modifié le ... :   /  /    
Description .. : Après mise à jour
Mots clefs ... : 
*****************************************************************}
procedure TOM_CALCULSALAIREMOY.OnAfterUpdateRecord ;
begin
  Inherited ;
  SetControlEnabled('METHODECAL',True);
end ;

procedure TOM_CALCULSALAIREMOY.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 05/06/2006
Modifié le ... :   /  /    
Description .. : Chargement de l'enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_CALCULSALAIREMOY.OnLoadRecord ;
begin
  Inherited ;
  if ds.state <> dsinsert then
begin
  // Gestion des prédéfinis
  LectureSeule := FALSE;

  if (Getfield('PCY_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
   end;

  if (Getfield('PCY_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
  end;

  if (Getfield('PCY_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), False);
  end;

  if (lectureseule = true) then Exit;

  SetControlEnabled('PCY_PREDEFINI', False);
  SetControlEnabled('PCY_CODECALSAL', False);
 end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 05/06/2006
Modifié le ... :   /  /    
Description .. : On change Field
Mots clefs ... : 
*****************************************************************}
procedure TOM_CALCULSALAIREMOY.OnChangeField ( F: TField ) ;
var
pred : string;
begin
  Inherited ;
  if DS.State in [DsEdit,DsInsert] then
    SetControlEnabled('METHODECAL',False)
  else
    SetControlEnabled('METHODECAL',True);

    if (F.FieldName = 'PCY_PREDEFINI') and (ds.State = DsInsert)  then
    begin

    Pred := GetField('PCY_PREDEFINI');

    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de code calcul prédéfini CEGID.', Ecran.Caption);
      Pred := 'DOS';
      SetControlProperty('PCY_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de code calcul prédéfini Standard.', Ecran.Caption);
      Pred := 'DOS';
      SetControlProperty('PCY_PREDEFINI', 'Value', Pred);
    end;
    if Pred <> GetField('PCY_PREDEFINI') then SetField('PCY_PREDEFINI', pred);
  end;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 05/06/2006
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOM_CALCULSALAIREMOY.OnArgument ( S: String ) ;
begin
  Inherited ;
  AccesPredefini('TOUS', CEG, STD, DOS);
end ;

procedure TOM_CALCULSALAIREMOY.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_CALCULSALAIREMOY.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_CALCULSALAIREMOY ] ) ; 
end.
