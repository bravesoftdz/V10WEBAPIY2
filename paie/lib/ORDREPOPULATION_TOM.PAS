{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : ORDREPOPULATION (ORDREPOPULATION)
                 Saisie des populations
Mots clefs ... : TOM;ORDREPOPULATION
*****************************************************************}
{PT1 22/02/07 NA Contrôle transversalité : affichage des valeurs du coefficient ou du libellé de la qualif selon le prédéfini
 PT2 22/02/07 NA Contrôle unicité des populations
 PT3 27/02/07 NA Contrôle de la suppression : habilitations, paramètres de paie, élement nationaux,zones libres salariés
 PT4 06/05/07 NA Si création, modif ou suppression d'une population : dévalide la population + contrôle validité
 PT5 10/07/07 NA Contrôle existence pour suppression : ajout  du prédéfini
 PT7 11/03/08 GGU FQ 15297 ne pas permettre de minuscule dans les codes alphanumériques
}

 
Unit ORDREPOPULATION_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     fe_main,
     HDB,
{$else}
     eFiche,
     maineagl,
{$ENDIF}
     sysutils,
     ComCtrls,
     HCtrls,
     HMsgBox,
     PgOutils,
     UTOM,
     P5Def,
     HTB97,
     pgpopuloutils,
     Hent1,
     UTobDebug,
     UTob ;

Type
  TOM_ORDREPOPULATION = Class (TOM)
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
 LectureSeule, CEG, STD, DOS, premfois : boolean;
 sauvcodepop, predavantmodif , Modemodif: string;
 nbident : integer;  // pt2
 btn,btn_detcrit : Ttoolbarbutton97;
 ListParam : Tstringlist;
 procedure Salariesaffectes(sender: TObject);
 procedure DetailCriteres(sender : TObject);
 procedure gestionpred;
 procedure initfiche;
 procedure initfichecreat;
 procedure LectureCmbBx(prompt_cmb: string; CmbBx: THMultiValComboBox);
 function  controlesalpop: boolean; // pt2
 procedure devalidepop;   // pt4
 end ;

Implementation


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /
Description .. : Création d'une population
*****************************************************************}
procedure TOM_ORDREPOPULATION.OnNewRecord ;
begin
  Inherited ;
  setcontrolenabled('PPC_PREDEFINI', false);
  setfocuscontrol('PPC_CODEPOP');
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /
Description .. :  Suppression Détail population
Mots clefs ... :
*****************************************************************}
procedure TOM_ORDREPOPULATION.OnDeleteRecord ;
var
Sql1, population,Sql2, sql3, sql4, Letitre1,lib : string;
suppop: Boolean;        // pt3

begin
  Inherited ;
  suppop:= true;
  lib := Getfield('PPC_LIBELLE');
  Letitre1 := ecran.caption + ' ' + lib;
  population := Getfield('PPC_POPULATION');

// Contrôle si la suppression est possible : pas de simulation existante pour ce détail pop

 Sql1 := 'Select PSR_CODESIMUL from RESULTSIMUL where ##PSR_PREDEFINI## PSR_POPULATION = "'+population+'"'; // pt5
  if ExisteSql(Sql1) then
  begin
   PGIBOX('Suppression impossible : Simulations existantes.', LeTitre1);
   suppop := false;
   Lasterror := 1;
   Exit;
  end;


  if suppop = true then
  begin
    Sql2 := 'Select PPR_CODEPOP from REGLESPOP where ##PPR_PREDEFINI## PPR_POPULATION = "'+population+'"';   // pt5
    if ExisteSql(Sql2) then
    begin             // si régles associées
      PGIBOX('Suppression impossible : il existe des règles IDR associées.', LeTitre1);
      suppop := false;
      lasterror := 1;
      Exit;
    end;
  end;
 // pt3
  if suppop = true then
  begin
    Sql3 := 'Select PHL_POPULATION from PGHABILITATIONS where ##PHL_PREDEFINI## PHL_POPULATION = "'+population+'"';   // pt5
    if ExisteSql(Sql3) then
    begin             // si habilitations pour cette population
      PGIBOX('Suppression impossible : il existe des habilitations associées.', LeTitre1);
      suppop := false;
      lasterror := 1;
      Exit;
    end;
  end;

  if suppop = true then
  begin
    sql4 := 'Select PGO_CODEASSOCIE FROM PGPARAMETRESASSOC WHERE ##PGO_PREDEFINI## PGO_CODEASSOCIE= "'+population+'" AND PGO_TYPEPARAMETRE = "POP"';  // pt5
     if ExisteSql(Sql4) then
     begin             // si paramètres associés pour cette population
      PGIBOX('Suppression impossible : il existe des paramètres de paie associés.', LeTitre1);
      suppop := false;
      lasterror := 1;
      Exit;
    end;
  end;

  if suppop = true then
  begin
    sql4 := 'Select PED_VALEURNIVEAU FROM ELTNATIONDOS WHERE ##PED_PREDEFINI## PED_VALEURNIVEAU = "'+population+'" AND PED_TYPENIVEAU = "POP"'; // pt5
     if ExisteSql(Sql4) then
     begin             // si éléments nationaux saisis au niveau de cette population
      PGIBOX('Suppression impossible : il existe des éléments nationaux saisis au niveau de cette population.', LeTitre1);
      suppop := false;
      lasterror := 1;
      Exit;
    end;
  end;

  if suppop = true then
  begin
    sql4 := 'Select PHD_POPULATION FROM PGHISTODETAIL WHERE ##PHD_PREDEFINI## PHD_POPULATION = "'+population+'"';  // pt5
     if ExisteSql(Sql4) then
     begin             // si zones libres salarié saisis pour cette population
      PGIBOX('Suppression impossible : il existe des zones libres salarié saisies au niveau de cette population.', LeTitre1);
      suppop := false;
      lasterror := 1;
      Exit;
    end;
  end;

  if suppop = true then
  begin
    sql4 := 'Select PFF_POPULATION FROM FORFAITFORM WHERE ##PFF_PREDEFINI## PFF_POPULATION = "'+population+'"'; // pt5
     if ExisteSql(Sql4) then
     begin             // si des fôrfaits prévisionnels ont été saisis pour cette population
      PGIBOX('Suppression impossible : il existe des forfaits prévisionnels saisis pour cette population.', LeTitre1);
      suppop := false;
      lasterror := 1;
      Exit;
    end;
  end;

  if suppop = true then
  begin
    sql4 := 'Select PFP_POPULATION FROM FRAISSALPLAF WHERE ##PFP_PREDEFINI## PFP_POPULATION = "'+population+'"'; // pt5
     if ExisteSql(Sql4) then
     begin             // si des plafonds de frais ont été  pour cette population
      PGIBOX('Suppression impossible : il existe des plafonds de frais saisis pour cette population.', LeTitre1);
      suppop := false;
      lasterror := 1;
      Exit;
    end;
  end;

  Modemodif := 'D';
  // fin pt3

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 11/04/2006
Modifié le ... :   /  /
Description .. : Controles de cohérence
Mots clefs ... :
*****************************************************************}
procedure TOM_ORDREPOPULATION.OnUpdateRecord ;
VAR
Sql3, population, codepop , pred, typepop, stypepop: string;
Q : Tquery;
salpopok : boolean;

begin
  Inherited ;
  SetField('PPC_CODEPOP', UpperCase(GetField('PPC_CODEPOP'))); //PT7
  SetField('PPC_POPULATION', UpperCase(GetField('PPC_POPULATION'))); //PT7
  codepop := Getfield('PPC_CODEPOP');
  population := Getfield('PPC_POPULATION');

  if (DS.State = dsinsert) then
  begin
  // controle que le code population est renseigné
    if codepop = '' then
    begin
      PGIBox('Le code population doit être renseigné. ', Ecran.caption);
      setfocuscontrol('PPC_CODEPOP');
      lasterror := 1;
      exit;
    end;
  // controle validité du code population
    Q:= Opensql('Select PPO_PREDEFINI, PPO_NODOSSIER, PPO_LIBELLE from CODEPOPULATION where ##PPO_PREDEFINI## PPO_CODEPOP = "'+codepop+'"',true); // pt5
    if Q.EOF then
    begin
      PGIBox('Le code population '+Codepop+ ' n''existe pas. ', Ecran.caption);
      setfocuscontrol('PPC_CODEPOP');
      lasterror := 1;
      exit;
    end;
    Pred := Q.findfield('PPO_PREDEFINI').asstring;
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
       PGIBox('Vous ne pouvez pas créer de population prédéfinie CEGID.', Ecran.caption);
       setfocuscontrol('PPC_CODEPOP');
       LastError := 1;
       exit;
      end;

    if (Pred = 'STD') and (STD = FALSE) then
    begin
       PGIBox('Vous ne pouvez pas créer de population prédéfinie Standard.', Ecran.caption);
       setfocuscontrol('PPC_CODEPOP');
       LastError := 1;
       exit;
    end;
  // la population doit être renseignée
    if (getfield('PPC_POPULATION') = '') then
    begin
      PGIBOX('Vous devez renseigner la population.', Ecran.caption);
      setfocuscontrol('PPC_POPULATION');
      Lasterror := 1;
      Exit;
    end;
  // La population doit être unique
    sql3 := 'select PPC_POPULATION from ORDREPOPULATION where ##PPC_PREDEFINI## PPC_POPULATION = "'+Population+'"';
    if ExisteSql(sql3) then
    begin
      PGIBOX('La population "'+population+'"  existe déjà', Ecran.caption);
      setfocuscontrol('PPC_POPULATION');
      Lasterror := 1;
      Exit;
    end;

    if  ((pred = 'STD') and (population[3] <> '1') and (population[3] <> '3')) then
    begin
      PGIBOX('La population doit se terminer par 1 ou 3.', Ecran.caption);
      setfocuscontrol('PPC_POPULATION');
      Lasterror := 1;
      exit;
    end;
    if ((pred = 'CEG') and (population[3] <> '0') and (population[3] <> '2') and (population[3] <> '4') and
    (population[3] <> '6') and (population[3] <> '8')) then
    begin
      PGIBOX('La population doit se terminer par 0,2,4,6 ou 8', Ecran.caption);
      setfocuscontrol('PPC_POPULATION');
      Lasterror := 1;
      exit;
    end;
    If ((pred = 'DOS') and (population[3] <> '5') and (population[3] <> '7') and (population[3] <> '9')) then
    begin
      PGIBOX('La population doit se terminer par 5,7 ou 9.', Ecran.caption);
      setfocuscontrol('PPC_POPULATION');
      Lasterror := 1;
      exit;
    end;

  end;
 // Contrôle qu'au moins une valeur de l'identifiant est renseigné
   if ((getfield('PPC_VALIDENT1') = '') and (getfield('PPC_VALIDENT2') = '') and (getfield('PPC_VALIDENT3') = '') and (getfield('PPC_VALIDENT4') = ''))
   then
   begin
     PGIBOX('Vous devez renseigner au moins une valeur identifiant.', Ecran.caption);
     setfocuscontrol('PPC_VALIDENT1');
     Lasterror := 1;
     Exit;
   end;

   // pt2
   // contrôle unicité de la population : recherche des salariés appartenant à ce type de population
    salpopok := controlesalpop;
    if salpopok = false  then   // pas unicité de la population
    begin
      PGIBOX('Population non valide : des salariés sont déjà affectés à une population, voir page anomalies.', Ecran.caption);
      //setfocuscontrol('PPC_VALIDENT1');
      (GetControl('PAGES') as TPageControl).ActivePageIndex := 1;
      Lasterror := 1;
      Exit;
   end;

    // fin pt2

  setcontrolenabled('B_SALAFFECTES', true);

  // recherche du prédéfini pour les types pop  Avant la validation
  typepop := getcontroltext('TYPEPOP');
  Stypepop := ReadTokenSt(Typepop);
  predavantmodif := GetPredefiniPopulation(stypepop);
  Modemodif := 'C';
 end;




{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/07/2006
Modifié le ... :   /  /
Description .. : Chargement
Mots clefs ... :
*****************************************************************}
procedure TOM_ORDREPOPULATION.OnLoadRecord ;

begin
  Inherited ;

  premfois := true;
  sauvcodepop := '';

  if ds.state = dsinsert  then initfichecreat   // initialisation de la fiche vierge
  else
  begin
  initfiche;      // initialisation de la fiche
  gestionpred;   // Gestion des prédéfinis
  setcontrolenabled('B_SALAFFECTES', true);
  end;


  if ds.state in [dsbrowse] then setcontrolenabled('B_SALAFFECTES', true);

end ;



{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /
Description .. : Affichage du libellé des identifiants +
Gestion des prédéfinis
Mots clefs ... :
*****************************************************************}
procedure TOM_ORDREPOPULATION.OnArgument ( S: String ) ;
begin
  Inherited ;

AccesPredefini('TOUS', CEG, STD, DOS);

Btn := TToolBarButton97(GetControl('B_SALAFFECTES'));
if btn <> nil
  then  Btn.Onclick := Salariesaffectes;

  btn_detcrit := TToolBarButton97(GetControl('BDetail'));
  if btn_detcrit <> nil then
    btn_detcrit.Onclick := DetailCriteres;
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/07/2006
Modifié le ... :   /  /
Description .. : saisie du code population
Mots clefs ... :
*****************************************************************}
procedure TOM_ORDREPOPULATION.OnChangeField(F: TField);
var
Q : TQuery;
pred, nodos, codepop : string;

begin
  inherited;

  if DS.State in [DsEdit,DsInsert] then
    SetControlEnabled('B_SALAFFECTES',False)
  else
    SetControlEnabled('B_SALAFFECTES',True);



  if (F.FieldName = 'PPC_CODEPOP') and (ds.State = DsInsert) and (premfois = false)  then
  begin
   codepop := Getfield('PPC_CODEPOP');

   if codepop <> sauvcodepop then
   begin
    sauvcodepop := codepop;

    // Contrôle que le code population soit saisi
    if codepop = '' then
    begin
      PGIBox('Vous devez renseigner le code population.', Ecran.caption);
      setfocuscontrol('PPC_CODEPOP');
      initfichecreat;
      exit;
    end;

    Q:= Opensql('Select PPO_PREDEFINI, PPO_NODOSSIER, PPO_LIBELLE from CODEPOPULATION where ##PPO_PREDEFINI## PPO_CODEPOP = "'+codepop+'"',true);
    if Q.EOF then
    begin
      PGIBox('Le code population '+Codepop+ ' n''existe pas. ', Ecran.caption);
      setfocuscontrol('PPC_CODEPOP');
      sauvcodepop := '';
      setfield('PPC_CODEPOP', '');
      initfichecreat;
      exit;
    end;


    if not Q.EOF then
    begin

      Pred := Q.findfield('PPO_PREDEFINI').asstring;
      nodos := Q.findfield('PPO_NODOSSIER').asstring;

      if (Pred = 'CEG') and (CEG = FALSE) then
      begin
       PGIBox('Vous ne pouvez pas créer de population prédéfinie CEGID.', Ecran.caption);
       setfocuscontrol('PPC_CODEPOP');
       sauvcodepop := '';
       setfield('PPC_CODEPOP', '');
       initfichecreat;
       LastError := 1;
       exit;
      end;

      if (Pred = 'STD') and (STD = FALSE) then
      begin
       PGIBox('Vous ne pouvez pas créer de population prédéfinie Standard.', Ecran.caption);
       setfocuscontrol('PPC_CODEPOP');
       sauvcodepop := '';
       setfield('PPC_CODEPOP', '');
       initfichecreat;
       LastError := 1;
       exit;
      end;
      setfield('PPC_PREDEFINI', pred);
      setfield('PPC_NODOSSIER', Nodos);
      setcontrolenabled('PPC_PREDEFINI', false);
      initfiche;
    end;
    ferme(Q);
   end;
  end;
  if (F.FieldName = 'PPC_CODEPOP') then premfois := false;

end;

{***********A.G.L.****************************************************
Auteur  ...... : NA
Créé le ...... : 11/04/2006
Modifié le ... :   /  /
Description .. : Initialisation de la fiche selon le code population
Mots clefs ... :
**********************************************************************}
procedure TOM_ORDREPOPULATION.Initfiche ;
var
st,Sti,lib, stplus  : string;
QQ : Tquery;
Tob_Ident,T1 : Tob;
i : integer;
begin

FreeAndNil(T1); nbident := 0;
st := 'SELECT PPO_NBIDENT, PPO_IDENT1, PPO_IDENT2, PPO_IDENT3, PPO_IDENT4,PPO_LIBELLE,PPO_TYPEPOP FROM CODEPOPULATION' +
' Where PPO_CODEPOP = "'+Getfield('PPC_CODEPOP')+'" and PPO_PREDEFINI = "'+Getfield('PPC_PREDEFINI')+'" and PPO_NODOSSIER = "'+Getfield('PPC_NODOSSIER')+'"';
QQ := OPENSQL(st, True);
if not QQ.Eof then
    Begin
    setcontroltext('TYPEPOP', QQ.findfield('PPO_TYPEPOP').asstring);
    Tob_Ident := Tob.Create('Les idents',nil,-1);
    Tob_Ident.LoadDetailDB('Les idents','','',QQ,False);
    T1 := Tob_Ident.FindFirst([''],[''],False);
    nbident := T1.getValue('PPO_NBIDENT');
    lib := T1.getValue('PPO_LIBELLE');
    SetControlProperty('LIBELLECODEPOP','Caption',lib);
    End;
Ferme(QQ);

if assigned(T1) then

   For i := 1 to 4 do
     Begin
     stplus := '';
     Sti := IntToStr(i);
     if i <= nbident then
       Begin
       Setcontrolvisible ('TPPC_VALIDENT'+Sti, True);
       Setcontrolvisible ('PPC_VALIDENT'+Sti, True);
       st := 'select PAI_LIBELLE,PAI_COLONNE, PAI_SUFFIX,PAI_PREFIX from PAIEPARIM where PAI_IDENT = '+IntToStr(T1.getValue('PPO_IDENT'+Sti))+'' ;
       QQ := OPENSQL(st, True);
       if not QQ.eof then
       SetControlProperty('TPPC_VALIDENT'+Sti,'Caption',QQ.Findfield('PAI_LIBELLE').Asstring);
       Setcontrolproperty('PPC_VALIDENT'+Sti,'Datatype',Rechdom('PGCHAMPTABL',T1.getValue('PPO_IDENT'+Sti),false));
       // si la donnée est le coefficient : la tablette est PGCOEFFICIENT et non pas PGLIBCOEFFICIENT
       if QQ.Findfield('PAI_COLONNE').asstring = 'COEFF' then
        Setcontrolproperty('PPC_VALIDENT'+Sti,'Datatype','PGCOEFFICIENT');

       // deb PT1
       if (QQ.findfield('PAI_COLONNE').asstring = 'COEFF') or (QQ.findfield('PAI_COLONNE').asstring = 'QUALIF')
       then
       begin
          stplus := '';
          if getfield('PPC_PREDEFINI') = 'CEG' then stplus :=  ' AND PMI_PREDEFINI = "CEG"' else
          if getfield('PPC_PREDEFINI') = 'STD' then stplus :=  ' AND PMI_PREDEFINI <> "DOS"';

          SetControlProperty('PPC_VALIDENT'+sti,'Plus',Stplus);
       end;
       // fin PT1

       // Recherche libellé TRAVAILN1,TRAVAILN2,TRAVAILN3,TRAVAILN4
       if QQ.Findfield('PAI_COLONNE').Asstring = 'TRAVN1' then
       VisibiliteChampSalarie ('1', GetControl('PPC_VALIDENT'+sti),GetControl('TPPC_VALIDENT'+sti));
       if QQ.Findfield('PAI_COLONNE').Asstring = 'TRAVN2' then
       VisibiliteChampSalarie ('2', GetControl('PPC_VALIDENT'+sti),GetControl('TPPC_VALIDENT'+sti));
       if QQ.Findfield('PAI_COLONNE').Asstring = 'TRAVN3' then
       VisibiliteChampSalarie ('3', GetControl('PPC_VALIDENT'+sti),GetControl('TPPC_VALIDENT'+sti));
        if QQ.Findfield('PAI_COLONNE').Asstring = 'TRAVN4' then
       VisibiliteChampSalarie ('4', GetControl('PPC_VALIDENT'+sti),GetControl('TPPC_VALIDENT'+sti));

       // Recherche libellé Code stat
       if QQ.Findfield('PAI_COLONNE').Asstring = 'CODSTAT' then
       Visibilitestat (GetControl('PPC_VALIDENT'+sti),GetControl('TPPC_VALIDENT'+sti));

       // Recherche libellé COMBO
        if QQ.Findfield('PAI_COLONNE').Asstring = 'LCMBX1' then
       VisibiliteChamplibresal ('1', GetControl('PPC_VALIDENT'+sti),GetControl('TPPC_VALIDENT'+sti));
       if QQ.Findfield('PAI_COLONNE').Asstring = 'LCMBX2' then
       VisibiliteChamplibresal ('2', GetControl('PPC_VALIDENT'+sti),GetControl('TPPC_VALIDENT'+sti));
       if QQ.Findfield('PAI_COLONNE').Asstring = 'LCMBX3' then
       VisibiliteChamplibresal ('3', GetControl('PPC_VALIDENT'+sti),GetControl('TPPC_VALIDENT'+sti));
        if QQ.Findfield('PAI_COLONNE').Asstring = 'LCMBX4' then
       VisibiliteChamplibresal ('4', GetControl('PPC_VALIDENT'+sti),GetControl('TPPC_VALIDENT'+sti));


       // pt2
       setcontroltext('NOMCHAMP'+sti, QQ.findfield('PAI_PREFIX').asstring + '_' + QQ.findfield('PAI_SUFFIX').asstring);
       // fin pt2

       ferme(QQ);
       End
    else
       Begin
       setcontroltext('NOMCHAMP'+sti, '');
       Setcontrolvisible ('TPPC_VALIDENT'+Sti, False);
       Setcontrolvisible ('PPC_VALIDENT'+Sti, False);
       End;
     End;
FreeAndNil(Tob_Ident);
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 11/04/2006
Modifié le ... :   /  /
Description .. : Initialisation de la fiche vierge
Mots clefs ... :
*****************************************************************}
procedure TOM_ORDREPOPULATION.Initfichecreat;
var
k : integer;
stk : string;

begin
SetControlEnabled('PPC_POPULATION', true);
SetControlEnabled('PPC_CODEPOP', true);
SetControlProperty('LIBELLECODEPOP','Caption','');
for k := 1 to 4 do
begin
Stk := IntToStr(k);
Setcontrolvisible ('TPPC_VALIDENT'+Stk, False);
Setcontrolvisible ('PPC_VALIDENT'+Stk, False);
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 11/04/2006
Modifié le ... :   /  /
Description .. : Gestion des prédéfinis
Mots clefs ... :
*****************************************************************}
procedure TOM_ORDREPOPULATION.gestionpred ;
begin
// Gestion des prédéfinis
  LectureSeule := FALSE;

  if GetField('PPC_PREDEFINI') = 'CEG' then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
  end;

  if GetField('PPC_PREDEFINI')= 'STD' then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
  end;

  if GetField('PPC_PREDEFINI') = 'DOS' then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), False);
  end ;

  // si modif et non lecture seule : code pop non saisissable
  setcontrolEnabled('PPC_PREDEFINI', False);
  if (ds.state <> dsinsert) and (lectureseule = false) then
  begin
   SetControlEnabled('PPC_POPULATION', False);
   SetControlEnabled('PPC_CODEPOP', False);
  end;
end;



{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 04/07/2006
Modifié le ... :   /  /
Description .. : Salairés affectés à cette population
Mots clefs ... :
*****************************************************************}
procedure TOM_ORDREPOPULATION.salariesaffectes;
Var
population, libellepop, st, codepop : string;
begin
population := Getfield('PPC_POPULATION');
libellepop := Getfield('PPC_LIBELLE');
codepop := Getfield('PPC_CODEPOP');
St := codepop + ';' + population +  ';' + libellepop;
// Salariés affectés à cette population
  AGLLanceFiche('PAY','SALARIEPOPUL_MUL','','',St);
end;

{***********A.G.L.***********************************************
Auteur  ...... : GHA
Créé le ...... : 16/01/2007
Modifié le ... : 19/01/2007
Description .. : Aperçu détaillé des critères sélectionnés
Mots clefs ... :
*****************************************************************}
procedure TOM_ORDREPOPULATION.DetailCriteres(sender: TObject);
var
    id_lib : integer;
    str_lib : variant;
    Cmbx : THMultiValComboBox;
begin
  ListParam := Tstringlist.Create;
  ListParam.Clear;

  for id_lib:= 1 to 4 do
  begin
    if GetControlVisible('TPPC_VALIDENT'+inttostr(id_lib)) then
    begin
     str_lib := GetControlText('TPPC_VALIDENT'+inttostr(id_lib)); //récupere le text du libellé.
     Cmbx := THMultiValComboBox(GetControl('PPC_VALIDENT'+inttostr(id_lib)));
     if Cmbx <> nil then
     begin
       // récupére les infos pour les alimenter dans une variable de type TstringList.
       LectureCmbBx(str_lib,Cmbx);
     end;
    end;
  end;
  Cmbx := nil;
  if not assigned(Cmbx) then
    cmbx.Free;

  AGLLanceFiche('PAY','DETAILPOPULATION','','',ListParam.Text);
  ListParam.free;
end;


procedure TOM_ORDREPOPULATION.OnClose ;
begin
  Inherited ;
  btn.Free;
  btn_detcrit.Free;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : GHA
Créé le ...... : 20/01/2007
Modifié le ... :   /  /    
Description .. : Lecture de la valeur des critères (pour visu des critères)
Mots clefs ... : 
*****************************************************************}
procedure TOM_ORDREPOPULATION.LectureCmbBx(prompt_cmb: string; CmbBx: THMultiValComboBox);
var
  mark,len : integer; //mark : position du séparateur / len : longueur de chaine
  str_item,source,separator,libelle : string;  //str_item : récupére la valeur de l'index de la combo box
begin
  source := CmbBx.Value;
  separator := ';';//CmbBx.Items.Delimiter;

  len := Length(source);
  while len <> 0 do
  begin
    mark := pos(separator,source);
    str_item := copy(source,1,mark-1);

    ListParam.Add(prompt_cmb); // critères
    ListParam.Add(str_item);   // codes
    libelle := CmbBx.Items.Strings[CmbBx.values.IndexOf(str_item)];
    ListParam.Add(libelle);    // libelles

    source := copy(source,mark+1,len-mark);
    len := Length(source);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 23/02/2007
Modifié le ... :   /  /    
Description .. : Controle unicité de la population
Mots clefs ... : 
*****************************************************************}
function TOM_ORDREPOPULATION.controlesalpop : boolean;  // pt2
var
st, typepop, st1, nomchamp, sti, sqlsal, salarie , population, libpop, nomsal , predefini: string;
tob_saltypepop, tids, tob_identsal, tidv, tob_valident , TD, tob_popsal: tob;
daterech : tdatetime;
i : integer;
pganomalies : Tlistbox;
QQ : Tquery;


begin
  result := true;
  pganomalies := TListbox(Getcontrol('PGANOMALIES'));
  pganomalies.Items.clear;
 // constitution de la clause where pour cette population
   Tob_Identsal := Tob.Create('Nom identifiants salarie', nil, -1);
   Tob_ValIdent := Tob.Create('Valeur identifiants', nil, -1);

  for i := 1 to nbident do
  begin
    sti := inttostr(i);
    Tids := Tob.Create('Identifiant', Tob_Identsal, -1);
    Tids.AddChampSup('NOMCHAMPS', False);
    nomchamp := getcontroltext('NOMCHAMP'+sti);
    Tids.PutValue('NOMCHAMPS', Nomchamp);
    Tids.AddChampSup('ORDREIDENT',False);
    Tids.PutValue('ORDREIDENT',sti);
  end;
    Tidv := Tob.Create('Identifiant', Tob_ValIdent, -1);
    Tidv.AddchampSupValeur('PPC_POPULATION', getcontroltext('PPC_POPULATION'));
    Tidv.AddchampSupValeur('PPC_LIBELLE', getcontroltext('PPC_LIBELLE'));
    Tidv.AddchampSupValeur('PPC_VALIDENT1', getcontroltext('PPC_VALIDENT1'));
    Tidv.AddchampSupValeur('PPC_VALIDENT2', getcontroltext('PPC_VALIDENT2'));
    Tidv.AddchampSupValeur('PPC_VALIDENT3', getcontroltext('PPC_VALIDENT3'));
    Tidv.AddchampSupValeur('PPC_VALIDENT4', getcontroltext('PPC_VALIDENT4'));

  st1 := '';

  if Assigned(tob_identsal) then st1 := PGConstitutionwhere(Tob_identsal, Tob_Valident);

  if assigned(Tob_valident) then FreeAndNil(Tob_valident);
  if assigned(Tob_Identsal) then FreeAndNil(Tob_identsal);



  daterech := date;
  // Recherche des salariés appartenant à ce type de population
  if st1 <> '' then
  sqlsal := 'select PSA_SALARIE from SALARIES where  '+st1+' and (PSA_DATESORTIE >="'+usdatetime(daterech)+'" or'+
  ' PSA_DATESORTIE <="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) and (PSA_DATEENTREE <= "'+
  usdatetime(daterech)+'")'
  else
  sqlsal := 'select PSA_SALARIE from SALARIES where (PSA_DATESORTIE >="'+usdatetime(daterech)+'" or'+
  ' PSA_DATESORTIE <="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) and (PSA_DATEENTREE <= "'+
  usdatetime(daterech)+'")';

  QQ:= Opensql(sqlsal, true);

    if tob_popsal <> nil then freeandnil(tob_popsal);
    Tob_popsal := Tob.Create('population_salarie',nil,-1);

     // lecture de chaque type de population  et constitution de la tob tob_popsal qui contient les salariés et leur population
     // pour chaque type de de population  et pour le prédéfini
    predefini := getfield('PPC_PREDEFINI');
    typepop := getcontroltext('TYPEPOP');
    St := ReadTokenSt(Typepop);
    While (St <> '') do
    Begin

     if st <> 'HAB' then   // ne pas controler l'unicité si population de type "habilitations"
     Begin

       tob_saltypepop := pgsalarietypepop(st,predefini,daterech);
       tob_popsal.dupliquer(tob_saltypepop,true,true,true);

       if tob_saltypepop <> nil then freeandnil(tob_saltypepop);

  // tobdebug(tob_popsal);
       QQ.first;
  // pour chaque salarié de la population, vérifie qu'il n'est pas déjà affecté à une population du même type.
       while not QQ.EOF do
       begin
          salarie := qq.findfield('PSA_SALARIE').asstring;
          TD := Tob_popsal.findfirst(['SALARIE'], [salarie], true);
          while  TD <> nil do
          begin  // un salarié appartient dèjà à une population ?
            population := TD.getvalue('POPULATION');
            libpop := TD.getvalue('LIBELLEPOP');
            nomsal := TD.getvalue('NOM')  + ' ' +  TD.getvalue('PRENOM');
            if population <>  getcontroltext('PPC_POPULATION') then
            begin
              Pganomalies.items.Add('le salarié ' + salarie + ' ' + nomsal + ' est déjà affecté '+
              'à la population' + ' ' +   population + ' ' + libpop + ' ' + ' pour le type' + ' ' + st);
              result := false;
              break;
            end;

            TD:= Tob_popsal.findNext(['SALARIE'], [salarie], true);
          End;
          QQ.next;
       End;
    End;
    St := ReadTokenSt(Typepop);

  End;

  ferme(QQ);

 if tob_popsal <> nil then freeandnil(tob_popsal);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 07/05/2007
Modifié le ... :   /  /    
Description .. : Dévalide les types de population
                 Affichage message info si changement de prédéfini pour les populations.
Mots clefs ... :
*****************************************************************}
procedure TOM_ORDREPOPULATION.Devalidepop;
// pt4
var
typepop,st , predef, sql, libpredef: string;
testidr, testpai, testfor : boolean;

begin
  testidr := false;
  testpai := false;
  
  typepop := getcontroltext('TYPEPOP');
  St := ReadTokenSt(Typepop);
  While (St <> '') do
  Begin
    setvalidepopulparamsoc(st, false);
    if st = 'IDR' then testidr := true;
    if st = 'PAI' then testpai := true;
    if st = 'FOR' then testfor := true;
    if st <> 'HAB' then
      if PGIAsk('Si vous avez terminé le paramétrage des populations de type '+ st
               +',#10#13vous devez lancer l''affectation des salariés.#10#13#10#13'
               +'Voulez vous effectuer l''affectation des salariés aux populations de type '+ st+ '?',Ecran.Caption)=mrYes  then
       AGLLanceFiche('PAY', 'AFFECTSALPOPUL', '', '',st);
    St := ReadTokenSt(Typepop);
  end;

   // si création d'une population avec un nouveau prédéfini et s'il existe des paramétrages zones libres
   // salarié ou éléments nationaux ou régles IDR avec les populations existantes ==> affichage message info
   predef := getcontroltext('PPC_PREDEFINI');
   typepop := getcontroltext('TYPEPOP');
   libpredef := 'CEGID';
   if predef = 'DOS' then libpredef := 'DOSSIER' else
   if predef = 'STD' then libpredef := 'STANDARD';



   if  predavantmodif <> predef  then
   begin
      PGIINFO('Les populations de type "'+typepop+'" utilisées seront celles prédéfinies "'+libpredef+'"', Ecran.Caption);
      Sql := 'Select PPR_CODEPOP from REGLESPOP WHERE ##PPR_PREDEFINI## PPR_PREDEFINI <> "'+predef+'"';

     if modemodif = 'C' then
     begin
        if (ExisteSql(Sql) and testidr)   then
        PGIINFO('Les règles IDR existantes ne seront plus utilisables.', Ecran.Caption);

        if testpai then
        begin
          sql := 'Select PGO_CODEASSOCIE FROM PGPARAMETRESASSOC WHERE ##PGO_PREDEFINI## PGO_TYPEPARAMETRE = "POP"';
          if ExisteSql(Sql) then
          PGIINFO('Les paramètres de paie associés aux populations ne seront plus valides.', Ecran.Caption);
          sql := 'Select PED_VALEURNIVEAU FROM ELTNATIONDOS WHERE ##PED_PREDEFINI## PED_TYPENIVEAU = "POP"';
          if ExisteSql(Sql) then
          PGIINFO('Les éléments nationaux saisis au niveau population ne seront plus valides.', Ecran.caption);
          sql := 'Select PHD_POPULATION FROM PGHISTODETAIL WHERE ##PHD_PREDEFINI## PHD_POPULATION <> ""';
          if ExisteSql(Sql) then
          PGIINFO('Les zones libres salarié saisies au niveau population ne seront plus valides.', Ecran.caption);
       end;
     end;
   end;
end;
// fin pt4

procedure TOM_ORDREPOPULATION.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_ORDREPOPULATION.OnAfterUpdateRecord ;
begin
  Inherited ;
   // pt4
  devalidepop;
  // fin pt4
end ;

procedure TOM_ORDREPOPULATION.OnAfterDeleteRecord ;
begin
  Inherited ;
   // pt4
  devalidepop;
  // fin pt4
end ;


Initialization
  registerclasses ( [ TOM_ORDREPOPULATION ] ) ;
end.
