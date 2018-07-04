{***********UNITE*************************************************
Auteur  ...... :  NA
Créé le ...... : 24/02/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : REGLESPOP (REGLESPOP)  REGLES IDR/POPULATION
Mots clefs ... : TOM;REGLESPOP
*****************************************************************
Pt1     :  17/04/2007    NA   Affichage des populations selon le prédéfini}
Unit REGLESPOP_TOM ;

Interface

Uses StdCtrls, 
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
{$else}
     eFiche, 
{$ENDIF}
     sysutils,
     ComCtrls,
     HCtrls,
     PgOutils,
     PgOutils2,
     HMsgBox,
     UTOM,
     Pgpopuloutils,
     UTob ;

Type
  TOM_REGLESPOP = Class (TOM)
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
    predpop : string;  // pt1
    procedure Gestionpred;
    end ;

Implementation

procedure TOM_REGLESPOP.OnNewRecord ;
begin
  Inherited ;
  if (CEG = TRUE) then
    SetField('PPR_PREDEFINI', 'CEG')
  else
    SetField('PPR_PREDEFINI', 'DOS');
    setfocuscontrol('PPR_POPULATION');
end ;

procedure TOM_REGLESPOP.OnDeleteRecord ;
begin
  Inherited ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /    
Description .. : Recherche du code population, prédéfini et dossier correspondant
                 au détail population saisi
Mots clefs ... : 
*****************************************************************}
procedure TOM_REGLESPOP.OnUpdateRecord ;
VAR
Pred,population,st,codepop : string;
QQ: TQUERY;


begin
  Inherited ;

  // Contrôle des saisies
  if Getfield('PPR_CODTABTUR') = '' then begin
  PGIBOX('Code table Turn OVER doit être renseigné.' , Ecran.caption);
  setfocuscontrol('PPR_CODTABTUR');
  Lasterror := 1;
  exit;
  end;

  if Getfield('PPR_CODTABREV') = '' then begin
  PGIBOX('Code table Revalorisation doit être renseigné.' , Ecran.caption);
  setfocuscontrol('PPR_CODTABREV');
  Lasterror := 1;
  exit;
  end;

  if Getfield('PPR_CODTABSUR') = '' then begin
  PGIBOX('Code table Survie doit être renseigné.' , Ecran.caption);
  setfocuscontrol('PPR_CODTABSUR');
  Lasterror := 1;
  exit;
  end;

  if Getfield('PPR_CODTABDRD') = '' then begin
  PGIBOX('Code table Droit Départ retraite "Départ volontaire" doit être renseigné.' , Ecran.caption);
  setfocuscontrol('PPR_CODTABDRD');
  Lasterror := 1;
  exit;
  end;

  if Getfield('PPR_AGEDEPD') = 0 then begin
  PGIBOX('L''age de départ en retraite "Départ volontaire" doit être renseigné.' , Ecran.caption);
  setfocuscontrol('PPR_AGEDEPD');
  Lasterror := 1;
  exit;
  end;

  if Getfield('PPR_CODTABDRR') = '' then begin
  PGIBOX('Code table Droit Départ retraite "Mise en retraite" doit être renseigné.' , Ecran.caption);
  setfocuscontrol('PPR_CODTABDRR');
  Lasterror := 1;
  exit;
  end;

  if Getfield('PPR_AGEDEPR') = 0 then begin
  PGIBOX('L''age de départ en retraite "Mise en retraite" doit être renseigné.' , Ecran.caption);
  setfocuscontrol('PPR_AGEDEPR');
  Lasterror := 1;
  exit;
  end;

  if Getfield('PPR_TXCHARGEPAT') = 0 then
  PGIBOX('Attention : Le taux de charge patronale n''est pas renseigné.' , Ecran.caption);

  // Recherche du code population

  population := getfield('PPR_POPULATION');
  st := 'SELECT PPC_CODEPOP FROM ORDREPOPULATION Where PPC_POPULATION = "'+population+'"';
  QQ := OPENSQL(st, True);
  if not QQ.Eof then codepop := QQ.findfield('PPC_CODEPOP').asstring;
  pred := getfield('PPR_PREDEFINI');
  if pred = 'DOS' then setfield('PPR_NODOSSIER', PgrendNodossier())  else
  setfield('PPR_NODOSSIER', '000000');

  setfield('PPR_CODEPOP', codepop);
  ferme(QQ);
end;

procedure TOM_REGLESPOP.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_REGLESPOP.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 19/09/2006
Modifié le ... :   /  /    
Description .. : Chargement de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOM_REGLESPOP.OnLoadRecord ;
var
 population, lib,  st: string;
 QQ: TQUERY;
Begin
Inherited ;
population := getfield('PPR_POPULATION');
st := 'SELECT PPC_LIBELLE FROM ORDREPOPULATION Where PPC_POPULATION = "'+population+'"';
QQ := OPENSQL(st, True);
if not QQ.Eof then lib := QQ.findfield('PPC_LIBELLE').Asstring;

ferme (QQ);
gestionpred;
Ecran.Caption := 'Règles IDR :' + population + ' ' + lib ;
UpdateCaption(Ecran);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 07/11/2006
Modifié le ... :   /  /    
Description .. : On change field
Mots clefs ... : 
*****************************************************************}
procedure TOM_REGLESPOP.OnChangeField ( F: TField ) ;
var
pred,st : string;
begin
  Inherited ;
   if (F.FieldName = 'PPR_PREDEFINI') and (ds.State = DsInsert)  then  // contrôle du prédéfini
  begin

    Pred := GetField('PPR_PREDEFINI');

    if Pred = '' then exit;

    if (Pred = 'CEG') then
    begin
      if  (CEG = FALSE) then
      begin // début pt1
        PGIBox('Vous ne pouvez pas créer une règle IDR prédéfinie CEGID.', Ecran.caption);
        Pred := 'DOS';
        SetControlProperty('PPR_PREDEFINI', 'Value', Pred);
      end
      else
      begin  // affichage des populations de type Cegid
        st := ' and PPC_PREDEFINI = "CEG"';
        setcontrolproperty('PPR_POPULATION', 'plus', st);
      end; // fin pt1
    end;

    if (Pred = 'STD') then
    begin
      if (STD = FALSE) then
      begin
        // début pt1
        PGIBox('Vous ne pouvez pas créer une règle IDR prédéfinie Standard.', Ecran.caption);
        Pred := 'DOS';
        SetControlProperty('PPR_PREDEFINI', 'Value', Pred);
      end
      else
      begin
        if predpop <> 'DOS' then pred:= predpop else pred := 'STD';
        st := ' and PPC_PREDEFINI = "'+pred+'"';
        setcontrolproperty('PPR_POPULATION', 'plus', st);
      end; // fin pt1
    end;
    // début pt1
    if (pred = 'DOS') then
    begin
       st := ' and PPC_PREDEFINI = "'+predpop+'"';
       setcontrolproperty('PPR_POPULATION', 'plus', st);
    end; // fin pt1

    if Pred <> GetField('PPR_PREDEFINI') then SetField('PPR_PREDEFINI', pred);
  end;
end ;

procedure TOM_REGLESPOP.OnArgument ( S: String ) ;
var
st, nodossier, pred : string;
Begin
  Inherited ;
  AccesPredefini('TOUS', CEG, STD, DOS);
  nodossier := PgrendNodossier();
  // recherche prédéfini des populations IDR
  pred := GetPredefiniPopulation('IDR');

  predpop := pred; // pt1
  st := ' and PPC_PREDEFINI = "'+pred+'"';
  setcontrolproperty('PPR_POPULATION', 'plus', st);

end ;

procedure TOM_REGLESPOP.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_REGLESPOP.OnCancelRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 07/11/2006
Modifié le ... :   /  /    
Description .. : Gestion des prédéfinis
Mots clefs ... : 
*****************************************************************}
procedure TOM_REGLESPOP.gestionpred ;
begin
// Gestion des prédéfinis
  LectureSeule := FALSE;


  if GetField('PPR_PREDEFINI') = 'CEG' then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
  end;

  if GetField('PPR_PREDEFINI')= 'STD' then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
  end;

  if GetField('PPR_PREDEFINI') = 'DOS' then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), False);
  end ;


  if lectureseule = true then
  begin
   setcontrolenabled('Binsert', false);
   setcontrolenabled('BDelete', false);
  end;


  // si modif et non lecture seule : zones clés non saisissables
  if (ds.state <> dsinsert) and (lectureseule = false) then begin
  SetControlEnabled('PPR_PREDEFINI', False);
  SetControlEnabled('PPR_POPULATION', False);
  end;

end;

Initialization
  registerclasses ( [ TOM_REGLESPOP ] ) ; 
end.
