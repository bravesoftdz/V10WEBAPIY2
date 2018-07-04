{***********UNITE*************************************************
Auteur  ...... :  NA
Créé le ...... : 20/01/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CODTAB (CODTAB)  Saisie des tables IDR
Mots clefs ... : TOM;CODTAB
*****************************************************************
PT1  : 10/04/2007 NA  Ajout code table dynamique + accés saisie table dynamique
PT3  : 17/04/2008 GGU V81 FQ 15361 Gestion uniformisée des zones libres - tables dynamiques

}

unit CODTAB_TOM;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fiche,
  fe_main,
{$ELSE}
  eFiche,
  maineagl,
  eFichList,
{$ENDIF}
  sysutils,
  ComCtrls,
  grids,
  HCtrls,
  HEnt1,
  HMsgBox,
  PgOutils,
  PgOutils2,
  UTOM,
  UTob,
  lookup,
  HDB,
  HTB97;


type
  TOM_CODTAB = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnAfterDeleteRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
  private
    Grille: THGrid;
    LeTitre: string;
    LectureSeule, CEG, STD, DOS, Err , codtabdynOK : boolean;
    procedure GrilleInsereLigne(sender: TObject);
    procedure GrilleDeleteLigne(sender: TObject);
    procedure GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure AlimGrille;
    procedure Grillemodif(sender: TObject);
    procedure Initgrille;
    procedure Gestionpred;
    procedure changetabledyn(sender:TObject); // pt1
    procedure tabledynelipsis(sender:TObject); // pt1
    procedure tabledynexit(sender:TObject); // pt1
    procedure Saisietabledyn(sender:TObject); // pt1
  end;

implementation

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /    
Description .. : Initialisation des zones QUATAB et PREDEFINI 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODTAB.OnNewRecord;
begin
  inherited;
  
  setfield('PTI_QUATAB', 'ANC');
  if (CEG = TRUE) then
    SetField('PTI_PREDEFINI', 'CEG')
  else
    SetField('PTI_PREDEFINI', 'DOS');
 end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /    
Description .. : Suppression des valeurs associées à la table
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODTAB.OnDeleteRecord;
var
sql1, codtab : string;

begin
  inherited;
  codtab := GetField('PTI_CODTAB');

  Sql1 := 'Select PPR_POPULATION  from REGLESPOP where  (PPR_CODTABTUR = "'+Codtab+'") or (PPR_CODTABREV = "'+Codtab+'")'+
  ' or (PPR_CODTABSUR = "'+Codtab+'") or (PPR_CODTABDRD = "'+Codtab+'") or (PPR_CODTABDRR = "'+Codtab+'")';
  if ExisteSql(Sql1) then
  begin
  PGIBOX('Suppression interdite : cette table est utilisée dans une règle IDR.' , Ecran.Caption);
  Lasterror := 1;
  exit;
  end;

  ExecuteSQL('DELETE FROM CODTABVALEUR WHERE PTS_PREDEFINI="' + GetField('PTI_PREDEFINI') +
    '" AND PTS_NODOSSIER ="' + GetField('PTI_NODOSSIER') + '" AND PTS_TYPTAB="' + GetField('PTI_TYPTAB') +
    '" AND PTS_CODTAB = "' + GetField('PTI_CODTAB') + '" ');
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /    
Description .. : Mise à jour de la table IDR : Suppression des valeurs 
Suite ........ : existantes et création des nouvelles valeurs saisies sur la 
Suite ........ : grille
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODTAB.OnUpdateRecord;
var
  Tob_Valeur, T1: TOB;
  val2,val3 : double;
  val1, i : Integer;
  typtab, codtab, predefini: string;
  tabledyn : Tcheckbox;
 begin

  inherited;


  // Contrôle du code table
  predefini := Getfield('PTI_PREDEFINI');
  codtab := Getfield('PTI_CODTAB');
  if codtab = '' then begin
  PGIBOX('Le code doit être renseigné.', Letitre);
  setfocuscontrol('PTI_CODTAB');
  Lasterror := 1;
  exit;
  end;

{  for k := 1 to 4 do
  begin
  if (codtab[k] in ['0'..'9'])  or (codtab[k] in ['A'..'Z']) or (codtab[k] in ['a'..'z']) then
  codeOK := true
  else
  codeOK:= false;
  if not codeOK then break;
  end;

  if not codeOK then
  begin
  PGIBOX('Le code doit comporter uniquement des chiffres ou des lettres.', Letitre);
  setfocuscontrol('PTI_CODTAB');
  Lasterror := 1;
  exit;
  end;    }



  if  ((predefini = 'STD') and (Codtab[4] <> '1') and (Codtab[4] <> '3'))
  then begin
  PGIBOX('Le code doit se terminer par 1 ou 3.', Letitre);
  setfocuscontrol('PTI_CODTAB');
  Lasterror := 1;
  exit;
  end;

  if ((predefini = 'CEG') and (Codtab[4] <> '0') and (Codtab[4] <> '2') and (codtab[4] <> '4') and
    (codtab[4] <> '6') and (codtab[4] <> '8'))
  then begin
  PGIBOX('Le code doit se terminer par 0,2,4,6 ou 8.', Letitre);
  setfocuscontrol('PTI_CODTAB');
  Lasterror := 1;
  exit;
  end;

  If ((predefini = 'DOS') and (Codtab[4] <> '5') and (Codtab[4] <> '7') and (Codtab[4] <> '9'))
  then begin
  PGIBOX('Le code doit se terminer par 5,7 ou 9.', Letitre);
  setfocuscontrol('PTI_CODTAB');
  Lasterror := 1;
  exit;
  end;

  if predefini = 'DOS' then setfield('PTI_NODOSSIER', PgrendNodossier())  else
  setfield('PTI_NODOSSIER', '000000');

  // contrôle que le code table saisi n'existe pas déjà pour une autre table si création
  if (ds.state = dsinsert)  then
  begin
     if existesql('select pti_codtab from CODTAB where pti_codtab = "'+codtab+'"') then
     begin
     PGIBOX('Ce code table existe déjà.', Letitre);
     setfocuscontrol('PTI_CODTAB');
     Lasterror := 1;
     exit;
     end;

  end;

  //Pt1 si table dynamique non OK ==> pas de validation
  tabledyn := Tcheckbox(Getcontrol('TABLEDYNAMIQUE'));
  if (tabledyn.checked = true)  then
  begin
    tabledynexit(nil);
    if (codtabdynOK = false) then
    begin
     Lasterror := 1;
     exit;
    end;
  end;
  // fin pt1

        // contrôle du type de table
      typtab := Getfield('PTI_TYPTAB');

      if typtab = '' then
      begin
        PGIBox('Le type n''est pas renseigné.', LeTitre);
        setfocuscontrol('PTI_TYPTAB');
        lasterror := 1;
        exit;
      end;

      // Err = true si 2 valeurs sont identiques
      if (err = true) then
      begin
        PGIBox('Vous ne pouvez pas renseigner 2 valeurs identiques.', LeTitre);
        lasterror := 1;
        exit;
      end;


      if Grille.rowcount - 1 >= 1 then
      begin
        Tob_Valeur := TOB.CREATE('Valeurtb', nil, -1);
        for i := 1 to grille.rowcount - 1 do
        begin

          if Grille.CellValues[0, i] <> '' then
          begin

          Grille.Cells[0, i] := FormatFloat(Grille.ColFormats[0], Valeur(Grille.Cells[0, i]));
          val1 := strtoint(Grille.CellValues[0, i]);


          T1 := TOB.CREATE('CODTABVALEUR', Tob_Valeur, -1);
          T1.PutValue('PTS_PREDEFINI', GetField('PTI_PREDEFINI'));
          T1.PutValue('PTS_NODOSSIER', GetField('PTI_NODOSSIER'));
          T1.PutValue('PTS_TYPTAB', typtab);
          T1.PutValue('PTS_CODTAB', GetField('PTI_CODTAB'));
          T1.PutValue('PTS_VALEUR', val1);

          if typtab = 'DRT' then
          begin
            Grille.Cells[1, i] := FormatFloat(Grille.ColFormats[1], Valeur(Grille.Cells[1, i]));
            val2 := strtofloat(Grille.CellValues[1,i]);
            T1.PutValue('PTS_NBMOIS', val2);
          end
          else
          if (typtab = 'REV') or (typtab = 'TUR') then
          begin
            Grille.Cells[1, i] := FormatFloat(Grille.ColFormats[1], Valeur(Grille.Cells[1, i]));
            val2 := strtofloat(Grille.CellValues[1,i]);
            T1.PutValue('PTS_TAUX', val2);
          end
          else
          if typtab = 'SUR' then
          begin
              Grille.Cells[1, i] := FormatFloat(Grille.ColFormats[1], Valeur(Grille.Cells[1, i]));
              val2 := strtoint(Grille.CellValues[1,i]);
              Grille.Cells[2, i] := FormatFloat(Grille.ColFormats[2], Valeur(Grille.Cells[2, i]));
              val3 := strtoint(Grille.CellValues[2,i]);
              T1.PutValue('PTS_EFFHOMME', val2);
              T1.PutValue('PTS_EFFFEMME', val3);
          end;
      end;
   end;
    // Suppression des valeurs saisies précédemment dans CODTABVALEUR
     ExecuteSQL('DELETE FROM CODTABVALEUR WHERE PTS_PREDEFINI="' + GetField('PTI_PREDEFINI') +
    '" AND PTS_NODOSSIER ="' + GetField('PTI_NODOSSIER') + '" AND PTS_TYPTAB="' + GetField('PTI_TYPTAB') +
    '" AND PTS_CODTAB = "' + GetField('PTI_CODTAB') + '" ');

    // Mise à jour des nouvelles valeurs saisies
    Tob_Valeur.InsertDB(nil, false);
    freeandnil(Tob_Valeur);
  end;

end;

procedure TOM_CODTAB.OnAfterUpdateRecord;
begin
  inherited;
end;

procedure TOM_CODTAB.OnAfterDeleteRecord;
begin
  inherited;

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /    
Description .. : Chargement des données : Initialisation et alimentation de la grille
Mots clefs ... :
*****************************************************************}
procedure TOM_CODTAB.OnLoadRecord;
var
tabledyn: Tcheckbox;// PT1
Q : TQuery; // PT1
Codetabledyn, typetable : string; //PT1
begin
  inherited;

  //PT1 : si code table dynamique renseigné pour les tables de type Droit ne pas afficher la grille 
  typetable:= getfield('PTI_TYPTAB');
  setcontroltext('NOMTABDYN', '');

  if typetable <> 'DRT' then
  begin
     setcontrolvisible('TABLEDYNAMIQUE', false);
     setcontrolvisible('PTI_CODTABLDYN', false);
  end
  else
  begin
    tabledyn := Tcheckbox(Getcontrol('TABLEDYNAMIQUE'));
    codetabledyn :=   getfield('PTI_CODTABLDYN');
    if codetabledyn = '' then
    tabledyn.checked := false
    else
    begin
      tabledyn.checked := true;
      setcontrolvisible('PTI_CODTABLDYN', true);
      setcontrolvisible('DETAILMINI', false);
      setcontrolvisible('PTI_QUATAB', false);
      setcontrolvisible('TPTI_QUATABIDR', false);
      Q:= opensql('SELECT PTE_LIBELLE from TABLEDIMENT where PTE_CODTABL = "'+codetabledyn+'"', true);
      if not Q.EOF then  setcontroltext('NOMTABDYN', Q.findfield('PTE_LIBELLE').asstring);
      ferme(Q);
    end;
  end;
  // fin PT1
  
  if Grille <> nil then
  begin
    initgrille;
    AlimGrille;
  end;
//  if lectureseule = false then Grille.rowcount := Grille.rowcount + 10;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /    
Description .. : Modification du type de table ou du qualifiant
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODTAB.OnChangeField(F: TField);
var
pred , libcol , typetabl, codtab, codtemp: string;
codenum : integer;
tabledyn : Tcheckbox; // pt1

begin
  inherited;
   // colle des 0 devant pour la saisie du code
  if F.fieldName = 'PTI_CODTAB' then
  begin
  codtab := Getfield('PTI_CODTAB');
   if codtab = '' then exit;
   if ((isnumeric(codtab)) and (codtab <> '   ')) then
   begin
   codenum := strtoint(trim(codtab));
   codtemp := collezerodevant(codenum,4);
   if codtemp <> codtab then setfield('PTI_CODTAB', codtemp);
   end;
  end;


  typetabl  := GetField('PTI_TYPTAB');
  
  if F.FieldName = 'PTI_TYPTAB' then
  begin
    initgrille;   // si modif type table ==> initialise la grille
    // pt1 pas de table dynamique si survie

    if typetabl <> 'DRT' then
    begin
       setcontrolvisible('TABLEDYNAMIQUE', false);
       setcontrolvisible('PTI_CODTABLDYN', false);
       setcontrolvisible('DETAILMINI', true);
    end
    else
    begin
       tabledyn := Tcheckbox(Getcontrol('TABLEDYNAMIQUE'));
       if ds.state = dsinsert then tabledyn.checked:= false;
       setcontrolvisible('TABLEDYNAMIQUE', true);
    end;
    // fin pt1
   end;

  if F.FieldName = 'PTI_QUATAB' then // si modif qualifiant ==> change libellé colonne
  begin
      libcol := RechDom('PGQUALIFTAB', GetField('PTI_QUATAB'), FALSE);
      if (Typetabl = 'REV') or (Typetabl = 'TUR') then
      libcol := libcol + ' ' + '(Seuil inférieur de la tranche)';

       grille.cells[0, 0] := libcol;
  end;


  if (F.FieldName = 'PTI_PREDEFINI') and (ds.State = DsInsert)  then  // contrôle du prédéfini
  begin

    Pred := GetField('PTI_PREDEFINI');

    if Pred = '' then exit;

    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de table IDR prédéfinie CEGID.', Letitre);
      Pred := 'DOS';
      SetControlProperty('PTI_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de table IDR prédéfinie Standard.', Letitre);
      Pred := 'DOS';
      SetControlProperty('PTI_PREDEFINI', 'Value', Pred);
    end;
    if Pred <> GetField('PTI_PREDEFINI') then SetField('PTI_PREDEFINI', pred);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /    
Description .. : Gestion des boutons : Insertion ligne ou suppression ligne
                 ou modification des données de la grille
Mots clefs ... :
*****************************************************************}
procedure TOM_CODTAB.OnArgument(S: string);
var
  Btn: TToolBarButton97;
  //deb pt1
  tabledyn : Tcheckbox;
  {$IFNDEF EAGLCLIENT}
  codetabledyn : THDBedit;
  {$ELSE}
  codetabledyn : THedit;
  {$ENDIF}
  // fin pt1

begin
  inherited;
  // Pt1 si coche table dynamique
  tabledyn := Tcheckbox(Getcontrol('TABLEDYNAMIQUE'));
  if tabledyn <> nil then tabledyn.onclick := changetabledyn;

  // si elipsis sur le code table dynamique
  {$IFNDEF EAGLCLIENT}
  codetabledyn := THDBEdit(GetControl('PTI_CODTABLDYN'));
  {$ELSE}
  codetabledyn := THEdit(GetControl('PTI_CODTABLDYN'));
  {$ENDIF}
   If codetabledyn <> Nil Then codetabledyn.OnElipsisClick := tabledynElipsis;
   if codetabledyn <> Nil Then codetabledyn.Onexit := tabledynexit;

   // Modif table dynamique
   Btn := TToolBarButton97(GetControl('B_TABLESDYNAM'));
   if Btn <> nil then Btn.OnClick := Saisietabledyn;

  // fin pt1

  Grille := THGrid(GetControl('DETAILMINI'));
  if Grille <> nil then
  begin
    Btn := TToolBarButton97(GetControl('B_INSERTION'));
    if Btn <> nil then Btn.OnClick := GrilleInsereLigne;
    Btn := TToolBarButton97(GetControl('B_SUPPRESSION'));
    if Btn <> nil then Btn.OnClick := GrilleDeleteLigne;
    Grille.OnEnter := Grillemodif;
    Grille.oncellExit := GrilleCellExit;
  end;
  LeTitre := 'Gestion des tables IDR';

  AccesPredefini('TOUS', CEG, STD, DOS);

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/05/2006
Modifié le ... :   /  /    
Description .. : Bouton Fermer
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODTAB.OnClose;
var
suite : integer;
begin
  inherited;
  if ((ds.state = dsedit) and (lectureseule = false))  then
  begin
    suite := Pgiask('Voulez-vous enregistrer les modifications ? ' ,Ecran.Caption);
    if (suite = mrYes) then TFFiche(Ecran).BValiderClick(Self);
  end;
end;

procedure TOM_CODTAB.OnCancelRecord;
begin
  inherited;
end;

procedure TOM_CODTAB.GrilleInsereLigne(sender: TObject);
begin
  Grille.Insertrow(Grille.Row);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /    
Description .. : Suppression d'une ligne de la grille
Mots clefs ... :
*****************************************************************}
procedure TOM_CODTAB.GrilleDeleteLigne(sender: TObject);
begin
  if Grille.Row <> 0 then Grille.deleterow(Grille.Row);
   if not (ds.state in [dsinsert, dsedit]) then ds.edit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /
Description .. : Alimentation de la grille avec les valeurs saisies 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODTAB.AlimGrille                                   ;
var
  TOB_VALTABLE: TOB;
  Valtable, typetabl: string;
  Qvaltable: Tquery;
  

begin
  Typetabl := GetField('PTI_TYPTAB');

  if (typetabl <> '') and (GetField('PTI_CODTAB') <> '') then
  begin
    Valtable := 'SELECT PTS_VALEUR,PTS_NBMOIS,PTS_TAUX,PTS_EFFHOMME,PTS_EFFFEMME from CODTABVALEUR' +
      ' where PTS_PREDEFINI = "' + GetField('PTI_PREDEFINI') + '" and PTS_NODOSSIER = "' + GetField('PTI_NODOSSIER') + '"' +
      ' and PTS_TYPTAB = "' + GetField('PTI_TYPTAB') + '" and PTS_CODTAB = "' + GetField('PTI_CODTAB') + '"';
    Qvaltable := OpenSQL(Valtable, True);

    if not Qvaltable.EOF  then
    begin
        TOB_VALTABLE := TOB.CREATE('Valeur table', nil, -1);
        TOB_VALTABLE.LoadDetailDB('CODTABVALEUR', '', '', Qvaltable, False);
        Ferme(Qvaltable);

        if typetabl = 'DRT' then
         TOB_VALTABLE.PutGridDetail(Grille, false, false, 'PTS_VALEUR;PTS_NBMOIS', True)
        else
         if (typetabl = 'REV') or (typetabl = 'TUR') then
         TOB_VALTABLE.PutGridDetail(Grille, false, false, 'PTS_VALEUR;PTS_TAUX', True)
         else
         if typetabl = 'SUR' then
         TOB_VALTABLE.PutGridDetail(Grille, false, false, 'PTS_VALEUR;PTS_EFFHOMME;PTS_EFFFEMME', True);

         freeandnil(TOB_VALTABLE);
         if lectureseule = false then Grille.rowcount := Grille.rowcount + 10;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 19/09/2006
Modifié le ... :   /  /    
Description .. : Grille modifiée
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODTAB.Grillemodif(sender: TObject);
begin
  if not (ds.state in [dsinsert, dsedit]) then
    ds.edit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /    
Description .. : Contrôle de la saisie en grille 
Mots clefs ... :
*****************************************************************}
procedure TOM_CODTAB.GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  i, lig: integer;
begin
  err := false;
  for i := 1 to THGrid(Sender).RowCount do
  begin
  if (THGrid(Sender).Cells[0, THGrid(Sender).Row]) <> '' then
  begin
    lig := THGrid(Sender).row;
    Grille.Cells[0, lig] := FormatFloat(Grille.ColFormats[0], Valeur(Grille.Cells[0, lig]));
    if (THGrid(Sender).Cells[0, THGrid(Sender).Row] = THGrid(Sender).Cells[0, i]) and (i <> THGrid(Sender).Row) then
    begin
        PGIBOX('Vous ne pouvez pas avoir deux valeurs identiques.', LeTitre);
        THGrid(Sender).Col := 0;
        THGrid(Sender).Row := ARow;
        err:= true;
        lasterror := 1;
        exit;
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /
Description .. : Initialisation de la grille selon le type de table IDR
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODTAB.Initgrille;
var
  Typetabl, libcol : string;
  i : integer;

begin
  setcontrolenabled('B_INSERTION', true);
  setcontrolenabled('B_SUPPRESSION', true);
  Grille.ColEditables[0] := true;
  Grille.ColEditables[1] := true;
  Grille.Coltypes[0] := 'R';
  Grille.Coltypes[1] := 'R';
  Grille.ColFormats[1] := '###0.000';
  Grille.ColFormats[0] := '###0';
  Grille.Colcount := 2;
  Grille.Colwidths[0] := 280;
  Grille.ColAligns[0] := taCenter;
  Grille.Colwidths[1] := 280;
  Grille.ColAligns[1] := taCenter;
  Typetabl := Getfield('PTI_TYPTAB');
  setcontrolenabled('PTI_QUATAB', true);
  if Typetabl = 'SUR' then
  begin
    Setfield('PTI_QUATAB', 'AGE');
   // setcontrolenabled('PTI_QUATAB', false);
    Grille.ColEditables[2] := true;
    Grille.Colcount := 3;
    Grille.Coltypes[2] := 'R';
    Grille.ColFormats[1] := '#####0';
    Grille.ColFormats[2] := '#####0';
    Grille.Colwidths[0] := 187;
    Grille.Colwidths[1] := 187;
    Grille.Colwidths[2] := 187;
    Grille.ColAligns[2] := taCenter;
    Grille.cells[1, 0] := 'Effectif Hommes';
    Grille.cells[2, 0] := 'Effectif Femmes';
  end
  else
    if Typetabl = 'DRT' then
    begin
      Grille.cells[1, 0] := 'Nombre de mois';
      setfield('PTI_QUATAB', 'ANC');
    //  setcontrolenabled('PTI_QUATAB', false);
    end
    else
      if (Typetabl = 'REV') or (Typetabl = 'TUR') then
        Grille.cells[1, 0] := 'Taux';

  libcol := RechDom('PGQUALIFTAB', GetField('PTI_QUATAB'), FALSE);
  if (Typetabl = 'REV') or (Typetabl = 'TUR') then
  libcol := libcol + ' ' + '(Seuil inférieur de la tranche)';

  grille.cells[0, 0] := libcol;
  
  // gestion des prédéfinis
  gestionpred;

  if (Typetabl = 'SUR') or (typetabl = 'DRT') then setcontrolenabled('PTI_QUATAB', false);

  if (lectureseule = false) and  (ds.State = DsInsert) then  begin
  grille.rowcount := 10;
   for i := 1 to Grille.rowcount do
   begin
    Grille.CellValues[0, i] := '';
    Grille.CellValues[1, i] := '';
    Grille.CellValues[2, i] := '';
   end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 14/04/2006
Modifié le ... :   /  /    
Description .. : Gestion des prédéfinis
Mots clefs ... :
*****************************************************************}
procedure TOM_CODTAB.gestionpred ;
begin
// Gestion des prédéfinis
  LectureSeule := FALSE;
  Grille := THGrid(GetControl('DETAILMINI'));

  if GetField('PTI_PREDEFINI') = 'CEG' then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
  end;

  if GetField('PTI_PREDEFINI')= 'STD' then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
  end;

  if GetField('PTI_PREDEFINI') = 'DOS' then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), False);
  end ;


  if lectureseule = true then
  begin
   setcontrolenabled('B_INSERTION', false);
   setcontrolenabled('B_SUPPRESSION', false);
  end;


  // si modif et non lecture seule : zones clés non saisissables
  if (ds.state <> dsinsert) and (lectureseule = false) then begin
  SetControlEnabled('PTI_PREDEFINI', False);
  SetControlEnabled('PTI_TYPTAB', False);
  SetControlEnabled('PTI_CODTAB', False);
  setcontrolenabled('PTI_QUATAB', false);
  end;


end;
//pt1
{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 10/04/2007
Modifié le ... :   /  /    
Description .. : Gestion de la Table dynamique
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODTAB.changetabledyn ;
var
tabledyn : Tcheckbox;

begin
tabledyn := Tcheckbox(Getcontrol('TABLEDYNAMIQUE'));

 if (tabledyn.checked) = true then
  begin
    setcontrolvisible('PTI_CODTABLDYN', true);
    setcontrolvisible('DETAILMINI', false);
    setcontrolvisible('PTI_QUATAB', false);
    setcontrolvisible('TPTI_QUATABIDR', false);
  end
  else
  begin
    setcontrolvisible('PTI_CODTABLDYN', false);
    setcontrolvisible('DETAILMINI', true);
    setcontrolvisible('PTI_QUATAB', true);
    setcontrolvisible('TPTI_QUATABIDR', true);
    setfield('PTI_CODTABLDYN', '');

  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 10/04/2007
Modifié le ... :   /  /    
Description .. : Elipsis code table dynamique
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODTAB.tabledynelipsis(Sender : TObject);
var
  pred, LookupWhere : string;
begin
  pred := getfield('PTI_PREDEFINI');
  LookupWhere := '';
  if pred = 'CEG' then
    LookupWhere := ' PTE_PREDEFINI = "CEG" ';
  if pred = 'STD' then
    LookupWhere := ' PTE_PREDEFINI <> "DOS" ';
  if pred = 'DOS' then
    LookupWhere := ' ##PTE_PREDEFINI## ';
  LookupWhere := LookupWhere
               + ' AND PTE_DTVALID = (SELECT MAX(Z.PTE_DTVALID) '
               + '                        FROM TABLEDIMENT Z '
//PT3               + '                       WHERE Z.PTE_CODTABL=TD.PTE_CODTABL '
                                         + ' WHERE ##Z.PTE_PREDEFINI## ' //PT3
                                         + ' AND TD.PTE_CODTABL = Z.PTE_CODTABL '
                                         + ' AND TD.PTE_PREDEFINI = Z.PTE_PREDEFINI '
                                         + ' AND TD.PTE_NODOSSIER = Z.PTE_NODOSSIER '
                                         + ' AND TD.PTE_NIVSAIS = Z.PTE_NIVSAIS '
                                         + ' AND TD.PTE_DTVALID<="'+UsDateTime(Date())+'")'
               + ' AND (PTE_CRITERE1 = 1500 OR PTE_CRITERE1 = 1501) '
               + ' AND (PTE_CRITERE2 = 0 or PTE_CRITERE2=1500 or PTE_CRITERE2=1501) ';
  LookupList(THEdit(Sender),'Tables dynamiques','TABLEDIMENT TD','DISTINCT PTE_CODTABL','PTE_LIBELLE ', LookupWhere, 'TD.PTE_CODTABL', True, -1);
end;
{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 10/04/2007
Modifié le ... :   /  /
Description .. : Sortie de l'elipsis code table dynamque
Mots clefs ... :
*****************************************************************}
procedure TOM_CODTAB.tabledynexit(Sender : TObject);
var
codetabledyn : string;
Q: Tquery;
critere1, critere2 : integer;
ok1, ok2 : boolean;

begin
 ok1 := false;
 ok2:= false;
 codtabdynOK := true;
 codetabledyn := getfield('PTI_CODTABLDYN');
 if codetabledyn <> '' then
 begin
   Q:= opensql('SELECT PTE_LIBELLE, PTE_CRITERE1,PTE_CRITERE2 from TABLEDIMENT where PTE_CODTABL = "'+codetabledyn+'"', true);
   if not Q.EOF then
   begin
      critere1 := Q.findfield('PTE_CRITERE1').Asinteger;
      critere2 := Q.findfield('PTE_CRITERE2').Asinteger;

          if (Critere1 = 1500) or (critere1 = 1501) then ok1 := true;
          if  (critere2 = 1500) or (critere2 = 1501) or (critere2 = 0) then  ok2:= true;
          if ok1 and ok2 then
          setcontroltext('NOMTABDYN', Q.findfield('PTE_LIBELLE').asstring)
          else
          begin
          codtabdynOK := false;
          PGIBOX('Cette table dynamique ne peut pas être utilisée pour les IDR.', Ecran.caption);
          setfocuscontrol('PTI_CODTABLDYN');
          end;
   end
   else
   begin
    setcontroltext('NOMTABDYN','');
    codtabdynOK := false;
    PGIBOX('Le code table dynamique n''existe pas.', Ecran.caption);
    setfocuscontrol('PTI_CODTABLDYN');
   end;
   ferme(Q);
 end
 else
 // code table non renseignée
 begin
    codtabdynOK := false;
    PGIBOX('Le code table dynamique doit être renseigné.', Ecran.caption);
    setfocuscontrol('PTI_CODTABLDYN');
 end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 14/05/2007
Modifié le ... :   /  /    
Description .. : Accès à la saisie des tables dynamiques
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODTAB.Saisietabledyn(Sender : TObject) ;
begin
 AglLanceFiche('PAY', 'PGMULTABLESDYNA', '', '', '');
end;
// fin pt1


initialization
  registerclasses([TOM_CODTAB]);
end.

