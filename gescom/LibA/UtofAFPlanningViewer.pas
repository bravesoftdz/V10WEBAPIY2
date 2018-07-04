{***********UNITE*************************************************
Auteur  ...... : CB
Créé le ...... : 22/05/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TOF_AFPLANNINGVIEWER ()
Mots clefs ... : TOF;TOF_AFPLANNINGVIEWER
*****************************************************************}
Unit UtofAFPlanningViewer;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,
{$IFNDEF EAGLCLIENT}
     db, dbtables, FE_Main,
{$ELSE}
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox, DicoAF, HTB97, hqry,  CalcOleGenericAff,
     UTOF, uTob, stat, AffaireUtil, UtilPlanning, 
     UtilTaches,UTofAfBaseCodeAffaire, paramsoc;

Type
  TOF_AFPLANNINGVIEWER = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    procedure FormKeyDown(SEnder: TObject; var Key: Word; ShIft: TShIftState);

    private
      fStAffaire : String;
      fTobViewer : Tob;

      // plan de charge en quantite
      procedure LoadData;
      procedure AddPlanifie;
      procedure AddRealise;
      procedure AddPlanifieCra;
      procedure AddEntete;
      procedure SuppNonPlanifie;
      procedure AddRAP;
      function AddClauseWhere(pStPrefixe : String) : String;
      function AddClauseFrom : String;
      procedure InitFiltre;

      // plan de charge en montant
      procedure LoadDataMnt;
      procedure AddPlanifieMnt;
      procedure AddRealiseMnt;
      procedure AddPlanifieCraMnt;
      procedure AddEnteteMnt;

  end;

  Function AFLanceFicheAFPlanningViewer(Lequel, Argument : String) : String;

const
	TexteMsgTache: array[1..1] of string 	= (
          {1}        'La ressource %s a du temps planifié et elle n''existe pas dans la tache %s de l''affaire %s !');

Implementation

Function AFLanceFicheAFPlanningViewer(Lequel, Argument : String) : String;
begin
  result := AGLLanceFiche('AFF','AFPLANNINGVIEWER','', '',Argument);
end;

procedure TOF_AFPLANNINGVIEWER.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFPLANNINGVIEWER.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFPLANNINGVIEWER.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFPLANNINGVIEWER.FormKeyDown(SEnder: TObject; var Key: Word; ShIft: TShIftState);
begin
  Inherited ;
  if (key = VK_ESCAPE) then close;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : Remplir la tob virtuelle avec les données du planning
                 et de l'activité
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGVIEWER.OnLoad ;
begin
  if fTobViewer = nil then
    fTobViewer := TOB.create('Tob ViewerMother', nil, -1)
  else
    fTobViewer.ClearDetail;
 
  if TCheckBox(getControl('CBPDCQUANTITE')).Checked then
    LoadData
  else
    LoadDataMnt;
  TFStat(Ecran).LaTOB :=  fTobViewer;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/06/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGVIEWER.LoadData;
var
  vSt     : String;
  vQR     : TQuery;
begin

  // Prevu, RAP par ressource
  vSt := 'SELECT ATA_AFFAIRE, ATA_AFFAIRE0, ATA_AFFAIRE1, ATA_AFFAIRE2,';
  vSt := vSt + ' ATA_AFFAIRE3, ATA_AVENANT, ATA_ARTICLE, GA_LIBELLE, ATA_FONCTION, AFO_LIBELLE,';
  vSt := vSt + ' ATR_NUMEROTACHE, ATR_RESSOURCE, ARS_LIBELLE,';
  vSt := vSt + ' ATR_QTEINITIALE AS Prevu,';
  vSt := vSt + ' 0.00 AS Realise, ';

  if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
    vSt := vSt + ' 0.00 AS AffectePA, 0.00 AS AffecteRess,'
  else
    vSt := vSt + ' 0.00 AS Affecte,';

  vSt := vSt + ' 0.00 AS ResteAPasser, ATR_QTEAPLANIFIER AS QteAPlanifier';
  vSt := vSt + ' ,GA_FAMILLENIV1 ';

  vSt := vSt + ' FROM AFFAIRE, TACHERESSOURCE, RESSOURCE, ARTICLE, TACHE ';
  vSt := vSt + ' LEFT OUTER JOIN FONCTION ON ATA_FONCTION = AFO_FONCTION ';

  vSt := vSt + AddClauseFrom;

  vSt := vSt + ' WHERE ATA_AFFAIRE = AFF_AFFAIRE';
  vSt := vSt + ' AND ATA_ARTICLE = GA_ARTICLE';

  vSt := vSt + ' AND ATA_TYPEARTICLE = "PRE" ';
  vSt := vSt + ' AND ATA_MODESAISIEPDC = "QUA" ';

  vSt := vSt + AddClauseWhere('ATA');

  vSt := vSt + ' AND ATA_AFFAIRE = ATR_AFFAIRE';
  vSt := vSt + ' AND ATA_NUMEROTACHE = ATR_NUMEROTACHE';
  vSt := vSt + ' AND ATR_RESSOURCE = ARS_RESSOURCE';

  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    if Not vQR.Eof then
      begin
        fTobViewer.LoadDetailDB('Mon TobViewer','','',vQr,False,True);
        AddPlanifie;
      end;

    // on affiche le realisé meme si il n'y a pas eu de planification
    AddRealise;
    if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then AddPlanifieCra;
    AddEntete;
    AddRAP;
    SuppNonPlanifie;

  Finally
    if vQR <> Nil then Ferme(vQR);
  End;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/06/2002
Modifié le ... :   /  /
Description .. : idem plan de charge en quantite mais en montant
Mots clefs ... : 
*****************************************************************}
procedure TOF_AFPLANNINGVIEWER.LoadDataMnt;
var
  vSt     : String;
  vQR     : TQuery;
begin

  // Prevu, RAP par ressource
  vSt := 'SELECT ATA_AFFAIRE, ATA_AFFAIRE0, ATA_AFFAIRE1, ATA_AFFAIRE2,';
  vSt := vSt + ' ATA_AFFAIRE3, ATA_AVENANT, ATA_ARTICLE, GA_LIBELLE, ';
  vSt := vSt + ' ATA_NUMEROTACHE, ';
  vSt := vSt + ' ATA_INITPTPR AS Prevu,';
  vSt := vSt + ' 0.00 AS Realise, ';

  if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
    vSt := vSt + ' 0.00 AS AffectePA, 0.00 AS AffecteRess,'
  else
    vSt := vSt + ' 0.00 AS Affecte,';

  vSt := vSt + ' 0.00 AS ResteAPasser, ATA_RAPPTPR AS QteAPlanifier ';
  vSt := vSt + ' ,GA_FAMILLENIV1 ';
  vSt := vSt + ' FROM AFFAIRE, TACHE, ARTICLE';

  vSt := vSt + AddClauseFrom;

  vSt := vSt + ' WHERE ATA_AFFAIRE = AFF_AFFAIRE';
  vSt := vSt + ' AND ATA_ARTICLE = GA_ARTICLE';

  vSt := vSt + ' AND ATA_TYPEARTICLE <> "PRE" ';
  vSt := vSt + ' AND ATA_MODESAISIEPDC = "MPR" ';
  vSt := vSt + AddClauseWhere('ATA');

  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    if Not vQR.Eof then
      begin
        fTobViewer.LoadDetailDB('Mon TobViewer','','',vQr,False,True);
        AddPlanifieMnt;
      end;

    // on affiche le realisé meme si il n'y a pas eu de planification
    AddRealiseMnt;
    if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then AddPlanifieCraMnt;
    AddEnteteMnt;
    AddRAP;
    SuppNonPlanifie;

  Finally
    if vQR <> Nil then Ferme(vQR);
  End;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : ajout des quantites planifiees
                 groupe par affaire, tache, article, fonction, ressource
                 remarque : tache = fonction = article
                 recherche les articles car on ne recupere que les articles
                 depuis le CRA
Mots clefs ... :
*****************************************************************}
procedure  TOF_AFPLANNINGVIEWER.AddPlanifie;
var
  i           : Integer;
  vSt         : String;
  vQR         : TQuery;
  vTob        : Tob;
  vTobAffecte : Tob;

Begin

  // planifie par ressource
  if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
    vSt := 'SELECT SUM(APL_QTEPLANIFIEE) AS AffectePA, '
  else
    vSt := 'SELECT SUM(APL_QTEPLANIFIEE) AS Affecte, ';

  vSt := vSt + ' APL_AFFAIRE AS ATA_AFFAIRE, APL_AFFAIRE0 AS ATA_AFFAIRE0, ';
  vSt := vSt + ' APL_AFFAIRE1 AS ATA_AFFAIRE1, APL_AFFAIRE2 AS ATA_AFFAIRE2,';
  vSt := vSt + ' APL_AFFAIRE3 AS ATA_AFFAIRE3, APL_AVENANT AS ATA_AVENANT, ';
  vSt := vSt + ' APL_NUMEROTACHE AS ATR_NUMEROTACHE, APL_ARTICLE AS ATA_ARTICLE,';
  vSt := vSt + ' GA_LIBELLE, APL_FONCTION AS ATA_FONCTION, APL_RESSOURCE AS ATR_RESSOURCE';
  vSt := vSt + ' FROM AFPLANNING, AFFAIRE, ARTICLE ';

  vSt := vSt + AddClauseFrom;

  vSt := vSt + ' WHERE APL_DATEDEBPLA >= "' + UsDateTime(DebutDeMois(now)) + '"';
  vSt := vSt + ' AND APL_ARTICLE = GA_ARTICLE';

  vSt := vSt + AddClauseWhere('APL');

  vSt := vSt + ' AND APL_RESSOURCE <> ""';
  vSt := vSt + ' AND APL_AFFAIRE = AFF_AFFAIRE';
  vSt := vSt + ' AND APL_TYPEARTICLE = "PRE" ';

  vSt := vSt + ' GROUP BY APL_AFFAIRE, APL_AFFAIRE0, APL_AFFAIRE1,';
  vSt := vSt + ' APL_AFFAIRE2, APL_AFFAIRE3, APL_AVENANT, ';
  vSt := vSt + ' APL_NUMEROTACHE, APL_ARTICLE, APL_FONCTION, APL_RESSOURCE, GA_LIBELLE';


  vTobAffecte := TOB.create('Tob Affecte', nil, -1);
  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    vTobAffecte.LoadDetailDB('Tob AffecteFille','','',vQR,False,True);

    for i := 0 to vTobAffecte.detail.count - 1 do
      begin
        vTob := fTobViewer.FindFirst(['ATA_AFFAIRE', 'ATA_ARTICLE', 'ATR_RESSOURCE', 'ATR_NUMEROTACHE'],
                             [vTobAffecte.Detail[i].GetValue('ATA_AFFAIRE'),
                              vTobAffecte.Detail[i].GetValue('ATA_ARTICLE'),
                              vTobAffecte.Detail[i].GetValue('ATR_RESSOURCE'),
                              vTobAffecte.Detail[i].GetValue('ATR_NUMEROTACHE')], true);
        if vTob <> nil then
        begin
          if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
            vTob.PutValue('AFFECTEPA', vTobAffecte.Detail[i].GetValue('AffectePA'))
          else
            vTob.PutValue('AFFECTE', vTobAffecte.Detail[i].GetValue('Affecte'))
        end
        // erreur
        else
          PGIBoxAF (format(TexteMsgTache[1],
                    [vTobAffecte.Detail[i].GetValue('ATR_RESSOURCE'),
                     vTobAffecte.Detail[i].GetValue('ATR_NUMEROTACHE'),
                     vTobAffecte.Detail[i].GetValue('ATA_AFFAIRE')]),'');
      end;
  Finally
    if vQR <> Nil then Ferme(vQR);
    vTobAffecte.Free;
  End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : ajout des quantites planifiees
                 groupe par affaire, tache, article, fonction, ressource
                 remarque : tache = fonction = article
                 recherche les articles car on ne recupere que les articles
                 depuis le CRA
Mots clefs ... :
*****************************************************************}
procedure  TOF_AFPLANNINGVIEWER.AddPlanifieMnt;
var
  i           : Integer;
  vSt         : String;
  vQR         : TQuery;
  vTob        : Tob;
  vTobAffecte : Tob;

Begin

  // planifie par ressource
  if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
    vSt := 'SELECT SUM(APL_INITPTPR) AS AffectePA, '
  else
    vSt := 'SELECT SUM(APL_INITPTPR) AS Affecte, ';

  vSt := vSt + ' APL_AFFAIRE AS ATA_AFFAIRE, APL_AFFAIRE0 AS ATA_AFFAIRE0, ';
  vSt := vSt + ' APL_AFFAIRE1 AS ATA_AFFAIRE1, APL_AFFAIRE2 AS ATA_AFFAIRE2,';
  vSt := vSt + ' APL_AFFAIRE3 AS ATA_AFFAIRE3, APL_AVENANT AS ATA_AVENANT, ';
  vSt := vSt + ' APL_NUMEROTACHE AS ATA_NUMEROTACHE, APL_ARTICLE AS ATA_ARTICLE,';
  vSt := vSt + ' GA_LIBELLE ';
  vSt := vSt + ' FROM AFPLANNING, AFFAIRE, ARTICLE ';

  vSt := vSt + AddClauseFrom;

  vSt := vSt + ' WHERE APL_DATEDEBPLA >= "' + UsDateTime(DebutDeMois(now)) + '"';
  vSt := vSt + ' AND APL_ARTICLE = GA_ARTICLE';

  vSt := vSt + AddClauseWhere('APL');

  vSt := vSt + ' AND APL_AFFAIRE = AFF_AFFAIRE';
  vSt := vSt + ' AND APL_TYPEARTICLE <> "PRE" ';

  vSt := vSt + ' GROUP BY APL_AFFAIRE, APL_AFFAIRE0, APL_AFFAIRE1,';
  vSt := vSt + ' APL_AFFAIRE2, APL_AFFAIRE3, APL_AVENANT, ';
  vSt := vSt + ' APL_NUMEROTACHE, APL_ARTICLE, GA_LIBELLE';


  vTobAffecte := TOB.create('Tob Affecte', nil, -1);
  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    vTobAffecte.LoadDetailDB('Tob AffecteFille','','',vQR,False,True);

    for i := 0 to vTobAffecte.detail.count - 1 do
      begin
        vTob := fTobViewer.FindFirst(['ATA_AFFAIRE', 'ATA_ARTICLE','ATA_NUMEROTACHE'],
                             [vTobAffecte.Detail[i].GetValue('ATA_AFFAIRE'),
                              vTobAffecte.Detail[i].GetValue('ATA_ARTICLE'),
                              vTobAffecte.Detail[i].GetValue('ATA_NUMEROTACHE')], true);
        if vTob <> nil then
        begin
          if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
            vTob.PutValue('AFFECTEPA', vTobAffecte.Detail[i].GetValue('AffectePA'))
          else
            vTob.PutValue('AFFECTE', vTobAffecte.Detail[i].GetValue('Affecte'))
        end;
      end;
  Finally
    if vQR <> Nil then Ferme(vQR);
    vTobAffecte.Free;
  End;
End;

procedure  TOF_AFPLANNINGVIEWER.AddRealise;
var
  vSt       : String;
  vQR       : TQuery;
  vTob      : Tob;
  vTobAct   : Tob;
  i         : Integer;

Begin

  // Realise
  vSt := vSt + ' SELECT ACT_AFFAIRE AS ATA_AFFAIRE, ACT_AFFAIRE0 AS ATA_AFFAIRE0, ';
  vSt := vSt + ' ACT_AFFAIRE1 AS ATA_AFFAIRE1, ACT_AFFAIRE2 AS ATA_AFFAIRE2,';
  vSt := vSt + ' ACT_AFFAIRE3 AS ATA_AFFAIRE3, ACT_AVENANT AS ATA_AVENANT, ';
  vSt := vSt + ' ACT_ARTICLE AS ATA_ARTICLE, GA_LIBELLE, AFO_FONCTION AS ATA_FONCTION, ';
  vSt := vSt + ' AFO_LIBELLE, 0 AS ATR_NUMEROTACHE, ACT_RESSOURCE AS ATR_RESSOURCE,';
  vSt := vSt + ' ARS_LIBELLE, 0.00 AS Prevu, SUM(ACT_QTE) AS Realise, ';

  if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
    vSt := vSt + ' 0.00 AS AffectePA, 0.00 AS AffecteRess'
  else
    vSt := vSt + ' 0.00 AS Affecte';

  vSt := vSt + ', 0.00 AS ResteAPasser, 0.00 AS QteAPlanifier ';
  vSt := vSt + ' ,GA_FAMILLENIV1 ';
  vSt := vSt + ' FROM AFFAIRE, RESSOURCE, ARTICLE, ACTIVITE ';
  vSt := vSt + ' LEFT OUTER JOIN FONCTION ON ACT_FONCTIONRES = AFO_FONCTION ';

  vSt := vSt + AddClauseFrom;

  vSt := vSt + ' WHERE ACT_ACTIVITEREPRIS <> "A"';
  vSt := vSt + ' AND ACT_ETATVISA = "VIS" ';
  vSt := vSt + ' AND ACT_ARTICLE = GA_ARTICLE';

  vSt := vSt + AddClauseWhere('ACT');

  vSt := vSt + ' AND ACT_AFFAIRE = AFF_AFFAIRE ';
  vSt := vSt + ' AND ACT_RESSOURCE = ARS_RESSOURCE ';
  vSt := vSt + ' AND ACT_TYPEARTICLE = "PRE" ';

  vSt := vSt + ' GROUP BY ACT_AFFAIRE, ACT_AFFAIRE0, ACT_AFFAIRE1,';
  vSt := vSt + ' ACT_AFFAIRE2, ACT_AFFAIRE3, ACT_AVENANT, ';
  vSt := vSt + ' ACT_ARTICLE, AFO_FONCTION, AFO_LIBELLE, ACT_RESSOURCE, ';
  vSt := vSt + ' ARS_LIBELLE, GA_LIBELLE,GA_FAMILLENIV1 ';

  vQr := nil;
  vTobAct := TOB.create('Tob ViewerMother', nil, -1);

  Try
    vQR := OpenSql(vSt,True);
    vTobAct.LoadDetailDB('Mon TobViewer','','',vQR,False,True);

    for i := vTobAct.Detail.count -1 downto 0 do
      begin
        vTob := fTobViewer.FindFirst(['ATA_AFFAIRE' , 'ATA_ARTICLE', 'ATR_RESSOURCE', 'ATA_FONCTION'],
                                     [vTobAct.Detail[i].GetValue('ATA_AFFAIRE'),
                                      vTobAct.Detail[i].GetValue('ATA_ARTICLE'),
                                      vTobAct.Detail[i].GetValue('ATR_RESSOURCE'),
                                      vTobAct.Detail[i].GetValue('ATA_FONCTION')], true);

        if vTob <> nil then
          vTob.PutValue('REALISE', vTobAct.Detail[i].GetValue('Realise'))

        // ajouter l'enregistrement dans la tob !
        else
          vTobAct.Detail[i].ChangeParent(fTobViewer, -1);

      end;
  Finally
    if vQR <> Nil then Ferme(vQR);
    vTobAct.Free;
  End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2003
Modifié le ... :   /  /
Description .. : idem Addrealise mais en montant
Mots clefs ... :
*****************************************************************}
procedure  TOF_AFPLANNINGVIEWER.AddRealiseMnt;
var
  vSt       : String;
  vQR       : TQuery;
  vTob      : Tob;
  vTobAct   : Tob;
  i         : Integer;

Begin

  // Realise
  vSt := vSt + ' SELECT ACT_AFFAIRE AS ATA_AFFAIRE, ACT_AFFAIRE0 AS ATA_AFFAIRE0, ';
  vSt := vSt + ' ACT_AFFAIRE1 AS ATA_AFFAIRE1, ACT_AFFAIRE2 AS ATA_AFFAIRE2,';
  vSt := vSt + ' ACT_AFFAIRE3 AS ATA_AFFAIRE3, ACT_AVENANT AS ATA_AVENANT, ';
  vSt := vSt + ' ACT_ARTICLE AS ATA_ARTICLE, GA_LIBELLE, 0 AS ATA_NUMEROTACHE, ';
  vSt := vSt + ' 0.00 AS Prevu, SUM(ACT_TOTPR) AS Realise, ';

  if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
    vSt := vSt + ' 0.00 AS AffectePA, 0.00 AS AffecteRess'
  else
    vSt := vSt + ' 0.00 AS Affecte';

  vSt := vSt + ' ,0.00 AS ResteAPasser, 0.00 AS QteAPlanifier ';
  vSt := vSt + ' ,GA_FAMILLENIV1 ';
  vSt := vSt + ' FROM AFFAIRE, ACTIVITE, ARTICLE ';

  vSt := vSt + AddClauseFrom;

  vSt := vSt + ' WHERE ACT_ACTIVITEREPRIS <> "A"';
  vSt := vSt + ' AND ACT_ETATVISA = "VIS" ';
  vSt := vSt + ' AND ACT_ARTICLE = GA_ARTICLE';

  vSt := vSt + AddClauseWhere('ACT');

  vSt := vSt + ' AND ACT_AFFAIRE = AFF_AFFAIRE ';
  vSt := vSt + ' AND ACT_TYPEARTICLE <> "PRE" ';

  vSt := vSt + ' GROUP BY ACT_AFFAIRE, ACT_AFFAIRE0, ACT_AFFAIRE1,';
  vSt := vSt + ' ACT_AFFAIRE2, ACT_AFFAIRE3, ACT_AVENANT, ';
  vSt := vSt + ' ACT_ARTICLE, GA_LIBELLE,GA_FAMILLENIV1 ';

  vQr := nil;
  vTobAct := TOB.create('Tob ViewerMother', nil, -1);

  Try
    vQR := OpenSql(vSt,True);
    vTobAct.LoadDetailDB('Mon TobViewer','','',vQR,False,True);

    for i := vTobAct.Detail.count -1 downto 0 do
      begin
        vTob := fTobViewer.FindFirst(['ATA_AFFAIRE', 'ATA_ARTICLE'],
                                     [vTobAct.Detail[i].GetValue('ATA_AFFAIRE'),
                                      vTobAct.Detail[i].GetValue('ATA_ARTICLE')], true);

        if vTob <> nil then
          vTob.PutValue('REALISE', vTobAct.Detail[i].GetValue('Realise'))

        // ajouter l'enregistrement dans la tob !
        else
          vTobAct.Detail[i].ChangeParent(fTobViewer, -1);
 
      end;
  Finally
    if vQR <> Nil then Ferme(vQR);
    vTobAct.Free;
  End;
end;

procedure  TOF_AFPLANNINGVIEWER.AddPlanifieCra;
var
  vSt       : String;
  vQR       : TQuery;
  vTob      : Tob;
  vTobAct   : Tob;
  i         : Integer;

Begin

  vSt := vSt + ' SELECT ACT_AFFAIRE AS ATA_AFFAIRE, ACT_AFFAIRE0 AS ATA_AFFAIRE0, ';
  vSt := vSt + ' ACT_AFFAIRE1 AS ATA_AFFAIRE1, ACT_AFFAIRE2 AS ATA_AFFAIRE2,';
  vSt := vSt + ' ACT_AFFAIRE3 AS ATA_AFFAIRE3, ACT_AVENANT AS ATA_AVENANT, ';
  vSt := vSt + ' ACT_ARTICLE AS ATA_ARTICLE, GA_LIBELLE, AFO_FONCTION AS ATA_FONCTION, ';
  vSt := vSt + ' AFO_LIBELLE, ACT_RESSOURCE AS ATR_RESSOURCE,';
  vSt := vSt + ' ARS_LIBELLE, 0.00 AS Prevu, 0.00 AS Realise,';

  if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
    vSt := vSt + ' 0.00 AS AffectePA, SUM(ACT_QTE) As AffecteRess '
  else
    vSt := vSt + ' 0.00 AS Affecte ';

  vSt := vSt + ' , 0.00 AS ResteAPasser, 0.00 AS QteAPlanifier ';
  vSt := vSt + ' ,GA_FAMILLENIV1 ';
  vSt := vSt + ' FROM AFFAIRE, RESSOURCE, ARTICLE, ACTIVITE ';
  vSt := vSt + ' LEFT OUTER JOIN FONCTION ON ACT_FONCTIONRES = AFO_FONCTION ';

  vSt := vSt + AddClauseFrom;

  vSt := vSt + ' WHERE ACT_ACTIVITEREPRIS = "A"';
  vSt := vSt + ' AND ACT_ETATVISA = "VIS" ';
  vSt := vSt + ' AND ACT_ARTICLE = GA_ARTICLE';

  vSt := vSt + AddClauseWhere('ACT');

  vSt := vSt + ' AND ACT_AFFAIRE = AFF_AFFAIRE ';
  vSt := vSt + ' AND ACT_RESSOURCE = ARS_RESSOURCE ';
  vSt := vSt + ' AND ACT_TYPEARTICLE = "PRE" ';

  vSt := vSt + ' GROUP BY ACT_AFFAIRE, ACT_AFFAIRE0, ACT_AFFAIRE1,';
  vSt := vSt + ' ACT_AFFAIRE2, ACT_AFFAIRE3, ACT_AVENANT, ';
  vSt := vSt + ' ACT_ARTICLE, AFO_FONCTION, AFO_LIBELLE, ACT_RESSOURCE, ';
  vSt := vSt + ' ARS_LIBELLE, GA_LIBELLE,GA_FAMILLENIV1 ';

  vTobAct := TOB.create('Tob ViewerMother', nil, -1);
  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    vTobAct.LoadDetailDB('Mon TobViewer','','',vQR,False,True);

    for i := vTobAct.Detail.count -1 downto 0 do
      begin
        vTob := fTobViewer.FindFirst(['ATA_AFFAIRE', 'ATA_ARTICLE', 'ATR_RESSOURCE', 'ATA_FONCTION'],
                                     [vTobAct.Detail[i].GetValue('ATA_AFFAIRE'),
                                      vTobAct.Detail[i].GetValue('ATA_ARTICLE'),
                                      vTobAct.Detail[i].GetValue('ATR_RESSOURCE'),
                                      vTobAct.Detail[i].GetValue('ATA_FONCTION')], true);
        if vTob <> nil then
          vTob.PutValue('AFFECTERESS', vTobAct.Detail[i].GetValue('AffecteRess'))

        // ajouter l'enregistrement dans la tob !
        else
          vTobAct.Detail[i].ChangeParent(fTobViewer, -1);

      end;
  Finally
    if vQR <> Nil then Ferme(vQR);
    vTobAct.Free;
  End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2003
Modifié le ... :   /  /
Description .. : idem AddPlanifieCra mais en montant
Mots clefs ... :
*****************************************************************}
procedure  TOF_AFPLANNINGVIEWER.AddPlanifieCraMnt;
var
  vSt       : String;
  vQR       : TQuery;
  vTob      : Tob;
  vTobAct   : Tob;
  i         : Integer;

Begin

  vSt := vSt + ' SELECT ACT_AFFAIRE AS ATA_AFFAIRE, ACT_AFFAIRE0 AS ATA_AFFAIRE0, ';
  vSt := vSt + ' ACT_AFFAIRE1 AS ATA_AFFAIRE1, ACT_AFFAIRE2 AS ATA_AFFAIRE2,';
  vSt := vSt + ' ACT_AFFAIRE3 AS ATA_AFFAIRE3, ACT_AVENANT AS ATA_AVENANT, ';
  vSt := vSt + ' ACT_ARTICLE AS ATA_ARTICLE, GA_LIBELLE, 0.00 AS Prevu, 0.00 AS Realise,';

  if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
    vSt := vSt + ' 0.00 AS AffectePA, SUM(ACT_TOTPR) As AffecteRess '
  else
    vSt := vSt + ' 0.00 AS Affecte ';

  vSt := vSt + ' , 0.00 AS ResteAPasser, 0.00 AS QteAPlanifier ';
  vSt := vSt + ' ,GA_FAMILLENIV1 ';
  vSt := vSt + ' FROM AFFAIRE, ACTIVITE, ARTICLE ';

  vSt := vSt + AddClauseFrom;

  vSt := vSt + ' WHERE ACT_ACTIVITEREPRIS = "A"';
  vSt := vSt + ' AND ACT_ETATVISA = "VIS" ';
  vSt := vSt + ' AND ACT_ARTICLE = GA_ARTICLE';

  vSt := vSt + AddClauseWhere('ACT');

  vSt := vSt + ' AND ACT_AFFAIRE = AFF_AFFAIRE ';
  vSt := vSt + ' AND ACT_TYPEARTICLE <> "PRE" ';

  vSt := vSt + ' GROUP BY ACT_AFFAIRE, ACT_AFFAIRE0, ACT_AFFAIRE1,';
  vSt := vSt + ' ACT_AFFAIRE2, ACT_AFFAIRE3, ACT_AVENANT, ';
  vSt := vSt + ' ACT_ARTICLE, GA_LIBELLE,GA_FAMILLENIV1 ';

  vTobAct := TOB.create('Tob ViewerMother', nil, -1);
  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    vTobAct.LoadDetailDB('Mon TobViewer','','',vQR,False,True);

    for i := vTobAct.Detail.count -1 downto 0 do
      begin
        vTob := fTobViewer.FindFirst(['ATA_AFFAIRE', 'ATA_ARTICLE', 'ATR_RESSOURCE', 'ATA_FONCTION'],
                                     [vTobAct.Detail[i].GetValue('ATA_AFFAIRE'),
                                      vTobAct.Detail[i].GetValue('ATA_ARTICLE'),
                                      vTobAct.Detail[i].GetValue('ATR_RESSOURCE'),
                                      vTobAct.Detail[i].GetValue('ATA_FONCTION')], true);
        if vTob <> nil then
          vTob.PutValue('AFFECTERESS', vTobAct.Detail[i].GetValue('AffecteRess'))

        // ajouter l'enregistrement dans la tob !
        else
          vTobAct.Detail[i].ChangeParent(fTobViewer, -1);

      end;
  Finally
    if vQR <> Nil then Ferme(vQR);
    vTobAct.Free;
  End;
end;

procedure TOF_AFPLANNINGVIEWER.AddRAP;
var
  i : Integer;

begin
  for i := 0 to fTobViewer.detail.count - 1 do
    begin
      if fTobViewer.detail[i].getValue('Prevu') = NULL then
        fTobViewer.detail[i].putValue('Prevu', 0);

      if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
      begin
        if fTobViewer.detail[i].getValue('AffectePA') = NULL then
          fTobViewer.detail[i].putValue('AffectePA', 0);
      end
      else
      begin
        if fTobViewer.detail[i].getValue('Affecte') = NULL then
          fTobViewer.detail[i].putValue('Affecte', 0);
      end;

      if fTobViewer.detail[i].getValue('Realise') = NULL then
        fTobViewer.detail[i].putValue('Realise', 0);

      if fTobViewer.detail[i].getValue('QteAPlanifier') = NULL then
        fTobViewer.detail[i].putValue('QteAPlanifier', 0);

      if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
        fTobViewer.detail[i].putValue('ResteAPasser',
                                      valeur(fTobViewer.detail[i].getValue('Prevu')) -
                                      valeur(fTobViewer.detail[i].getValue('AffectePA')) -
                                      valeur(fTobViewer.detail[i].getValue('Realise')) +
                                      valeur(fTobViewer.detail[i].getValue('QteAPlanifier')))
      else
        fTobViewer.detail[i].putValue('ResteAPasser',
                                      valeur(fTobViewer.detail[i].getValue('Prevu')) -
                                      valeur(fTobViewer.detail[i].getValue('Affecte')) -
                                      valeur(fTobViewer.detail[i].getValue('Realise')) +
                                      valeur(fTobViewer.detail[i].getValue('QteAPlanifier')))

    end;
end;
  
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : ajout des quantites planifiees en entete
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGVIEWER.AddEntete;
                                          
var
  i           : Integer;
  vSt         : String;
  vQR         : TQuery;
  vTob        : Tob;
  vTobAffecte : Tob;
                               
Begin

  vSt := vSt + ' SELECT APL_AFFAIRE AS ATA_AFFAIRE, APL_AFFAIRE0 AS ATA_AFFAIRE0, ';
  vSt := vSt + ' APL_AFFAIRE1 AS ATA_AFFAIRE1, APL_AFFAIRE2 AS ATA_AFFAIRE2,';
  vSt := vSt + ' APL_AFFAIRE3 AS ATA_AFFAIRE3, APL_AVENANT AS ATA_AVENANT, ';
  vSt := vSt + ' APL_ARTICLE AS ATA_ARTICLE, GA_LIBELLE, APL_FONCTION AS ATA_FONCTION, ';
  vSt := vSt + ' AFO_LIBELLE, APL_NUMEROTACHE AS ATR_NUMEROTACHE, ';
  vSt := vSt + ' "" AS ATR_RESSOURCE,';
  vSt := vSt + ' "" AS ARS_LIBELLE, 0.00 AS Prevu, 0.00 AS Realise,';

  if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
    vSt := vSt + ' SUM(APL_QTEPLANIFIEE) AS AffectePA, 0.00 As AffecteRess'
  else
     vSt := vSt + ' SUM(APL_QTEPLANIFIEE) AS Affecte ';

  vSt := vSt + ' , 0.00 AS ResteAPasser, 0.00 AS QteAPlanifier ';
  vSt := vSt + ' ,GA_FAMILLENIV1 ';
  vSt := vSt + ' FROM AFFAIRE, ARTICLE, AFPLANNING ';
  vSt := vSt + ' LEFT OUTER JOIN FONCTION ON APL_FONCTION = AFO_FONCTION  ';
                                                 
  vSt := vSt + AddClauseFrom;                       

  vSt := vSt + ' WHERE APL_DATEDEBPLA >= "' + UsDateTime(DebutDeMois(now)) + '"';
  vSt := vSt + ' AND APL_ARTICLE = GA_ARTICLE';

  vSt := vSt + AddClauseWhere('APL');

  vSt := vSt + ' AND APL_RESSOURCE = ""';
  vSt := vSt + ' AND APL_AFFAIRE = AFF_AFFAIRE';
  vSt := vSt + ' AND APL_TYPEARTICLE = "PRE" ';

  vSt := vSt + ' GROUP BY APL_AFFAIRE, APL_AFFAIRE0, APL_AFFAIRE1,';
  vSt := vSt + ' APL_AFFAIRE2, APL_AFFAIRE3, APL_AVENANT, ';
  vSt := vSt + ' APL_NUMEROTACHE, APL_ARTICLE, APL_FONCTION, APL_RESSOURCE, ';
  vSt := vSt + ' AFO_LIBELLE, GA_LIBELLE,GA_FAMILLENIV1 ';

  vTobAffecte := TOB.create('Tob Affecte', nil, -1);
  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    vTobAffecte.LoadDetailDB('Mon TobViewer','','',vQR,False,True);
                                             

    for i := vTobAffecte.Detail.count -1 downto 0 do
      begin
        vTob := fTobViewer.FindFirst(['ATA_AFFAIRE', 'ATA_ARTICLE', 'ATR_NUMEROTACHE'],
                             [vTobAffecte.Detail[i].GetValue('ATA_AFFAIRE'),
                              vTobAffecte.Detail[i].GetValue('ATA_ARTICLE'),
                              vTobAffecte.Detail[i].GetValue('ATR_NUMEROTACHE')], true);
        if vTob = nil then
          vTobAffecte.Detail[i].ChangeParent(fTobViewer, -1);
      end;

  Finally
    if vQR <> Nil then Ferme(vQR);
    vTobAffecte.Free;
  End;  
End;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2003
Modifié le ... :   /  /
Description .. : idem AddEntete mais en montant
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGVIEWER.AddEnteteMnt;

var
  i           : Integer;
  vSt         : String;
  vQR         : TQuery;
  vTob        : Tob;
  vTobAffecte : Tob;

Begin

  vSt := vSt + ' SELECT APL_AFFAIRE AS ATA_AFFAIRE, APL_AFFAIRE0 AS ATA_AFFAIRE0, ';
  vSt := vSt + ' APL_AFFAIRE1 AS ATA_AFFAIRE1, APL_AFFAIRE2 AS ATA_AFFAIRE2,';
  vSt := vSt + ' APL_AFFAIRE3 AS ATA_AFFAIRE3, APL_AVENANT AS ATA_AVENANT, ';
  vSt := vSt + ' APL_ARTICLE AS ATA_ARTICLE, GA_LIBELLE, ';
  vSt := vSt + ' APL_NUMEROTACHE AS ATA_NUMEROTACHE, ';
  vSt := vSt + ' 0.00 AS Prevu, 0.00 AS Realise,';
                                            
  if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
    vSt := vSt + ' SUM(APL_INITPTPR) AS AffectePA, 0.00 As AffecteRess'
  else
     vSt := vSt + ' SUM(APL_INITPTPR) AS Affecte ';

  vSt := vSt + ' , 0.00 AS ResteAPasser, 0.00 AS QteAPlanifier ';
  vSt := vSt + ' ,GA_FAMILLENIV1 ';
  vSt := vSt + ' FROM AFFAIRE, AFPLANNING, ARTICLE ';

  vSt := vSt + AddClauseFrom;

  vSt := vSt + ' WHERE APL_DATEDEBPLA >= "' + UsDateTime(DebutDeMois(now)) + '"';
  vSt := vSt + ' AND APL_ARTICLE = GA_ARTICLE';

  vSt := vSt + AddClauseWhere('APL');

  vSt := vSt + ' AND APL_RESSOURCE = ""';
  vSt := vSt + ' AND APL_AFFAIRE = AFF_AFFAIRE';
  vSt := vSt + ' AND APL_TYPEARTICLE <> "PRE" ';

  vSt := vSt + ' GROUP BY APL_AFFAIRE, APL_AFFAIRE0, APL_AFFAIRE1,';
  vSt := vSt + ' APL_AFFAIRE2, APL_AFFAIRE3, APL_AVENANT, ';
  vSt := vSt + ' APL_NUMEROTACHE, APL_ARTICLE, ';
  vSt := vSt + ' GA_LIBELLE,GA_FAMILLENIV1 ';//, ARS_LIBELLE ';

  vTobAffecte := TOB.create('Tob Affecte', nil, -1);
  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    vTobAffecte.LoadDetailDB('Mon TobViewer','','',vQR,False,True);
                             

    for i := vTobAffecte.Detail.count -1 downto 0 do
      begin
        vTob := fTobViewer.FindFirst(['ATA_AFFAIRE', 'ATA_ARTICLE', 'ATA_NUMEROTACHE'],
                             [vTobAffecte.Detail[i].GetValue('ATA_AFFAIRE'),
                              vTobAffecte.Detail[i].GetValue('ATA_ARTICLE'),
                              vTobAffecte.Detail[i].GetValue('ATA_NUMEROTACHE')], true);
        if vTob = nil then
          vTobAffecte.Detail[i].ChangeParent(fTobViewer, -1);
      end;

  Finally
    if vQR <> Nil then Ferme(vQR);
    vTobAffecte.Free;
  End;
End;

procedure TOF_AFPLANNINGVIEWER.OnArgument (S : String ) ;
var
  vStTmp    : String;
  vStChamp  : String;
  vStValeur : String;

Begin

  Inherited;

  // traitement des arguments
  vStTmp:= (Trim(ReadTokenSt(S)));
  While (vStTmp <>'') do
    Begin
      DecodeArgument(vStTmp, vStChamp, vStValeur);
      If vStChamp = 'AFFAIRE' then
        begin
          fStAffaire := vStValeur;
          //Gerer le filtre
          InitFiltre;
        end
      else if vStChamp = 'MONTANT' then
        TCheckBox(GetControl('CBPDCQUANTITE')).Checked := False;
            
      vStTmp:= (Trim(ReadTokenSt(S)));
    End;


  SetControlVisible('PAVANCE', false);
{$IFDEF EAGLCLIENT}
  setControlVisible('BIMPRIMER', false);
{$ENDIF}
                            
end;

procedure TOF_AFPLANNINGVIEWER.InitFiltre;
var
  vStAff0, vStAff1  : String;
  vStAff2, vStAff3  : String;
  vStAvenant        : String;
  vStTiers          : String;

Begin
  TFStat(ecran).CritereVisible := false;

  CodeAffaireDecoupe(fStAffaire, vStAff0, vStAff1, vStAff2,
                     vStAff3,vStAvenant, taModif, false);

  // C.B prise en compte des affaires modeles
  TeststCleAffaire(fStAffaire,vStAff0,vStAff1,vStAff2,vStAff3,vStAvenant,vStTiers, False, True, False, False);

  SetControlText('ATA_AFFAIRE',fStAffaire);
  SetControlText('ATA_AFFAIRE1',vStAff1);
  SetControlText('ATA_AFFAIRE2',vStAff2);
  SetControlText('ATA_AFFAIRE3',vStAff3);
  SetControlText('ATA_AVENANT',vStAvenant);
  SetControlText('ATA_TIERS', vStTiers);

  SetControlEnabled('ATA_AFFAIRE1', False);
  SetControlEnabled('ATA_AFFAIRE2', False);
  SetControlEnabled('ATA_AFFAIRE3', False);
  SetControlEnabled('ATA_AVENANT', False);
  SetControlEnabled('ATA_TIERS', False);
  SetControlEnabled('BEFFACEAFF1', False);
  SetControlEnabled('BSELECTAFF1', False);

  TToolBarButton97(GetControl('BCHERCHE')).Click;

end;

procedure TOF_AFPLANNINGVIEWER.OnClose;
begin
  Inherited;
  if assigned(fTobViewer) then
    begin
      fTobViewer.Free;
      fTobViewer := nil;
      TFStat(Ecran).LaTOB := nil;
    end;
end;
 
procedure TOF_AFPLANNINGVIEWER.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
Begin
  Aff   := THEdit(GetControl('ATA_AFFAIRE'));
  Aff1  := THEdit(GetControl('ATA_AFFAIRE1'));
  Aff2  := THEdit(GetControl('ATA_AFFAIRE2'));
  Aff3  := THEdit(GetControl('ATA_AFFAIRE3'));
  Aff4  := THEdit(GetControl('ATA_AVENANT'));
  Tiers := THEdit(GetControl('ATA_TIERS'));
End;

function TOF_AFPLANNINGVIEWER.AddClauseWhere(pStPrefixe : String) : String;
var
  vStParam : String;

begin
  result := copy(RecupWhereCritere(TPageControl(GetControl('PAGES'))), 6, length(RecupWhereCritere(TPageControl(GetControl('PAGES')))));
  if result <> '' then result := ' AND ' + result;

  vStParam := GetControlText('YTC_TABLELIBRETIERS1');
  if vStParam <> '' then result := result + ' AND YTC_TIERS = AFF_TIERS ';

  if (pStPrefixe <> 'ATA') then ReplaceSubStr(result, 'ATA_', pStPrefixe + '_');
end;

function TOF_AFPLANNINGVIEWER.AddClauseFrom : String;
var
  vStParam : String;
begin

  vStParam := GetControlText('YTC_TABLELIBRETIERS1');
  if vStParam <> '' then result := result + ', TIERSCOMPL ';
                                           
// table article ajoutée dans tous les cas
//  vStParam := GetControlText('GA_FAMILLENIV1');
//  if vStParam <> '' then result := result + ', ARTICLE ';

end;                            

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 17/09/2002
Modifié le ... :
Description .. : suppression des lignes ayant toutes les quantités à zéro
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGVIEWER.SuppNonPlanifie;
var
  i : Integer;

begin
  for i := fTobViewer.Detail.count -1 downto 0 do
  begin
    if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
    begin                            
      if (fTobViewer.detail[i].GetValue('Prevu') = 0) and
         (fTobViewer.detail[i].GetValue('AffectePA') = 0) and
         (fTobViewer.detail[i].GetValue('AffecteRess') = 0) and
         (fTobViewer.detail[i].GetValue('ResteAPasser') = 0) and
         (fTobViewer.detail[i].GetValue('QteAPlanifier') = 0) then
      begin
        fTobViewer.detail[i].Free;
      end;
    end
    else
    begin
      if (fTobViewer.detail[i].GetValue('Prevu') = 0) and
         (fTobViewer.detail[i].GetValue('Affecte') = 0) and
         (fTobViewer.detail[i].GetValue('ResteAPasser') = 0) and
         (fTobViewer.detail[i].GetValue('QteAPlanifier') = 0) then
      begin
        fTobViewer.detail[i].Free;
      end;
    end;
  end;
end;

Initialization
  registerclasses ([TOF_AFPLANNINGVIEWER]) ;
end.



