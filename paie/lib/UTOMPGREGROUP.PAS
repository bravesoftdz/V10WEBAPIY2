{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 28/09/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PGREGROUP (PGREGROUP)
Mots clefs ... : TOM;PGREGROUP
*****************************************************************}
Unit UTOMPGREGROUP ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList,
     fe_main,
{$else}
     eFiche,
     maineagl,
     eFichList, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM,
     PgOutils,
     PgOutils2,
     htb97,
     UTob ;

Type
  TOM_PGREGROUP = Class (TOM)
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
    procedure OnLoadAlerte               ; override ;
    private
    LectureSeule, CEG, STD, DOS : boolean;
    Typeregroupement : string;
    procedure BDetailclick(sender: Tobject);
    end ;


Implementation

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 28/09/2007
Modifié le ... :   /  /    
Description .. : Création
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGREGROUP.OnNewRecord ;
begin
  Inherited ;
    SetControlEnabled('B_DETAILREGROUP',False);
  if (CEG = TRUE) then
    SetField('PRB_PREDEFINI', 'CEG')
  else
    SetField('PRB_PREDEFINI', 'DOS');
end ;


procedure TOM_PGREGROUP.OnDeleteRecord ;
begin
  Inherited ;
  ExecuteSql('Delete from PGREGROUPDET where ##POT_PREDEFINI## POT_CODEREGROUP ="'+Getfield('PRB_CODEREGROUP')+'"');
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 28/09/2007
Modifié le ... :   /  /    
Description .. : Mise à jour
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGREGROUP.OnUpdateRecord ;
var predefini : string;
begin
  Inherited ;
  predefini := Getfield('PRB_PREDEFINI');
  if predefini = 'DOS' then setfield('PRB_NODOSSIER', PgrendNodossier())  else
  setfield('PRB_NODOSSIER', '000000');

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 28/09/2007
Modifié le ... :   /  /    
Description .. : Après Mise à jour
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGREGROUP.OnAfterUpdateRecord ;
begin
  Inherited ;
   SetControlEnabled('B_DETAILREGROUP',True);
end ;

procedure TOM_PGREGROUP.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 28/09/2007
Modifié le ... :   /  /    
Description .. : Chargement
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGREGROUP.OnLoadRecord ;
begin
  Inherited ;
    if ds.state <> dsinsert then
begin
  // Gestion des prédéfinis
  LectureSeule := FALSE;

  if (Getfield('PRB_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
   end;

  if (Getfield('PRB_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
  end;

  if (Getfield('PRB_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), False);
  end;

  if (lectureseule = true) then Exit;
  SetControlEnabled('PRB_PREDEFINI', False);
  SetControlEnabled('PRB_CODEREGROUP', False);
 end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 28/09/2007
Modifié le ... :   /  /    
Description .. : Modif champ
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGREGROUP.OnChangeField ( F: TField ) ;
var pred : string;
begin
  Inherited ;
   if DS.State in [DsEdit,DsInsert] then
    SetControlEnabled('B_DETAILREGROUP',False)
  else
    SetControlEnabled('B_DETAILREGROUP',True);

    if (F.FieldName = 'PRBREDEFINI') and (ds.State = DsInsert)  then
    begin

    Pred := GetField('PRB_PREDEFINI');

    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de code calcul prédéfini CEGID.', Ecran.Caption);
      Pred := 'DOS';
      SetControlProperty('PRB_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de code calcul prédéfini Standard.', Ecran.Caption);
      Pred := 'DOS';
      SetControlProperty('PRB_PREDEFINI', 'Value', Pred);
    end;
    if Pred <> GetField('PRB_PREDEFINI') then SetField('PRB_PREDEFINI', pred);
  end;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 28/09/2007
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGREGROUP.OnArgument ( S: String ) ;
var
 BDETAIL : TToolBarButton97;
 St , mode : string;
begin
  Inherited ;
   mode := readtokenst(S);
   typeregroupement := readtokenst(S);
   AccesPredefini('TOUS', CEG, STD, DOS);

   st := 'AND CO_ABREGE LIKE "%'+typeregroupement+'%"';
   SetControlProperty('PRB_THEMEREGROUP','Plus',St);
   if typeregroupement = 'MAB' then st := 'AND CO_ABREGE = "M"' else st := 'AND CO_ABREGE = "R"';
   SetControlProperty('PRB_TYPEREGROUP','Plus',St);

   BDETAIL := TToolBarButton97(GetControl('B_DETAILREGROUP'));
   if BDETAIL <> nil then BDETAIL.OnClick := BDetailClick;
end ;

procedure TOM_PGREGROUP.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_PGREGROUP.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_PGREGROUP.OnLoadAlerte;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 28/09/2007
Modifié le ... :   /  /    
Description .. : Détail du regroupement
Mots clefs ... : 
*****************************************************************}
procedure TOM_PGREGROUP.BDetailClick(sender: Tobject);
begin

   AGLLanceFiche('PAY', 'PGREGROUPDET', '','',typeregroupement +';'+ Getfield('PRB_THEMEREGROUP') +';'+
   GetField('PRB_CODEREGROUP')+';'+GetField('PRB_LIBELLE') + ';' + GetField('PRB_PREDEFINI')+ ';' + GetField('PRB_NODOSSIER'));

end;

Initialization
  registerclasses ( [ TOM_PGREGROUP ] ) ; 
end.
