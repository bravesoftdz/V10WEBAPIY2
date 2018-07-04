{***********UNITE*************************************************
Auteur  ...... : CB
Créé le ...... : 22/05/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TOF_AFPLANNINGRES ()
Mots clefs ... : TOF;TOF_AFPLANNINGRES
*****************************************************************}
Unit UtofAFPlanningRes;

Interface

Uses StdCtrls,
     Controls,
     Classes,       
     Windows,
{$IFNDEF EAGLCLIENT}
     db,                         
     dbtables,
     FE_Main,
{$ELSE}
     MaineAGL,     
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox, DicoAF, HTB97, hqry,CalcOleGenericAff,
     UTOF, uTob, stat, AffaireUtil, utilPlanning,
     UtilTaches,UTofAfBaseCodeAffaire, paramsoc;

Type
  TOF_AFPLANNINGRES = Class (TOF_AFBASECODEAFFAIRE)
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

      procedure CBPDCQUANTITEClick(Sender: TObject);

      // plan de charge en quantite
      procedure LoadData;
      function AddClauseWhere(pStPrefixe : String) : String;
      function AddClauseFrom : String;
      procedure InitFiltre;
      procedure SuppNonPlanifie(pTobLigne : Tob);
      procedure AddRessourceCRA(pTobViewer : Tob);

      // plan de charge en montant
      procedure LoadDataMnt;
  end;

  Function AFLanceFicheAFPlanningRes(Lequel, Argument : String) : String;

const
	TexteMsgTache: array[1..1] of string 	= (
          {1}        'La ressource %s n''existe pas dans la tache %s de l''affaire %s !');

Implementation

Function AFLanceFicheAFPlanningRes(Lequel, Argument : String) : String;
begin
  result := AGLLanceFiche('AFF','AFPLANNINGRES','', '',Argument);
end;

procedure TOF_AFPLANNINGRES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFPLANNINGRES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFPLANNINGRES.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFPLANNINGRES.FormKeyDown(SEnder: TObject; var Key: Word; ShIft: TShIftState);
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
procedure TOF_AFPLANNINGRES.OnLoad ;
begin
  if fTobViewer = nil then
    fTobViewer := TOB.create('TobViewerMother', nil, -1)
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
procedure TOF_AFPLANNINGRES.LoadData;
var
  vSt       : String;
  vQR       : TQuery;
  i         : Integer;
  vTobLigne : Tob;

begin

  vSt := 'SELECT ATA_AFFAIRE, ATA_AFFAIRE0, ATA_AFFAIRE1, ATA_AFFAIRE2,';
  vSt := vSt + ' ATA_AFFAIRE3, ATA_AVENANT, ATA_ARTICLE, GA_LIBELLE, ATA_FONCTION, AFO_LIBELLE,';
  vSt := vSt + ' ATR_NUMEROTACHE AS ATA_NUMEROTACHE, ATR_RESSOURCE AS ATA_RESSOURCE, ARS_LIBELLE, ';
  vSt := vSt + ' GA_FAMILLENIV1, AFF_LIBELLE, AFF_TIERS, ';
  vSt := vSt + ' 0.00 AS PC_MOIS, 0.00 AS PC_MOIS1, 0.00 AS PC_MOIS2, 0.00 AS PC_MOIS3, ';
  vSt := vSt + ' 0.00 AS PC_MOIS4, 0.00 AS PC_MOIS5, 0.00 AS PC_MOIS6, 0.00 AS PC_MOIS7, ';
  vSt := vSt + ' 0.00 AS PC_MOIS8, 0.00 AS PC_MOIS9, 0.00 AS PC_MOIS10, 0.00 AS PC_MOIS11, ';
  vSt := vSt + ' 0.00 AS PC_MOIS12, 0.00 AS PC_MOIS13, 0.00 AS PC_MOIS14, 0.00 AS PC_MOIS15, ';
  vSt := vSt + ' 0.00 AS PC_MOIS16, 0.00 AS PC_MOIS17 ';

  if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
  begin
    vSt := vSt + ' ,0.00 AS PC_MOISCRA, 0.00 AS PC_MOIS1CRA, 0.00 AS PC_MOIS2CRA, ';
    vSt := vSt + ' 0.00 AS PC_MOIS3CRA, 0.00 AS PC_MOIS4CRA, 0.00 AS PC_MOIS5CRA, ';
    vSt := vSt + ' 0.00 AS PC_MOIS6CRA, 0.00 AS PC_MOIS7CRA, 0.00 AS PC_MOIS8CRA, ';
    vSt := vSt + ' 0.00 AS PC_MOIS9CRA, 0.00 AS PC_MOIS10CRA, 0.00 AS PC_MOIS11CRA, ';
    vSt := vSt + ' 0.00 AS PC_MOIS12CRA, 0.00 AS PC_MOIS13CRA, 0.00 AS PC_MOIS14CRA, ';
    vSt := vSt + ' 0.00 AS PC_MOIS15CRA, 0.00 AS PC_MOIS16CRA, 0.00 AS PC_MOIS17CRA ';
  end;

  vSt := vSt + ' FROM AFFAIRE, TACHE, TACHERESSOURCE, RESSOURCE, FONCTION, ARTICLE ';

  vSt := vSt + AddClauseFrom;

  vSt := vSt + ' WHERE ATA_AFFAIRE = AFF_AFFAIRE';
  vSt := vSt + ' AND ATA_TYPEARTICLE = "PRE" ';
  vSt := vSt + ' AND ATA_ARTICLE = GA_ARTICLE ';

  vSt := vSt + AddClauseWhere('ATA');

  vSt := vSt + ' AND ATA_AFFAIRE = ATR_AFFAIRE';
  vSt := vSt + ' AND ATA_NUMEROTACHE = ATR_NUMEROTACHE';
  vSt := vSt + ' AND ATR_RESSOURCE = ARS_RESSOURCE';
  vSt := vSt + ' AND ATA_FONCTION = AFO_FONCTION';

  vSt := vSt + ' GROUP BY ATA_AFFAIRE, ATA_AFFAIRE0, ATA_AFFAIRE1,';
  vSt := vSt + ' ATA_AFFAIRE2, ATA_AFFAIRE3, ATA_AVENANT,';
  vSt := vSt + ' ATR_NUMEROTACHE, ATA_ARTICLE, ATA_FONCTION, AFO_LIBELLE, ';
  vSt := vSt + ' ATR_RESSOURCE, ARS_LIBELLE, GA_LIBELLE, GA_FAMILLENIV1, ';
  vSt := vSt + ' AFF_LIBELLE, AFF_TIERS ';
                                        
  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    if Not vQR.Eof then
      begin
         fTobViewer.LoadDetailDB('MONTOBVIEWER','','',vQr,False,True);

        // Ajout dans la liste des données du CRA pour les ressources non planifiées
        if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then AddRessourceCRA(fTobViewer);

        // pour chaque ressource
        for i := 0 to fTobViewer.Detail.Count - 1 do
          begin
            vTobLigne := fTobViewer.Detail[i];
            LoadMois(vTobLigne, vTobLigne, False, vTobLigne.Getvalue('ATA_AFFAIRE'), 'ATA', 'APL_QTEPLANIFIEE');

            if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
              LoadMoisCra(vTobLigne, vTobLigne.Getvalue('ATA_AFFAIRE'));

          end;
        // suppression des lignes
        SuppNonPlanifie(fTobViewer);
      end
    else
      begin
        // Ajout dans la liste des données du CRA pour les ressources non planifiées
        if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then AddRessourceCRA(fTobViewer);

        // pour chaque ressource
        for i := 0 to fTobViewer.Detail.Count - 1 do
          begin
            vTobLigne := fTobViewer.Detail[i];
            LoadMois(vTobLigne, vTobLigne, False, vTobLigne.Getvalue('ATA_AFFAIRE'), 'ATA', 'APL_QTEPLANIFIEE');

            if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
              LoadMoisCra(vTobLigne, vTobLigne.Getvalue('ATA_AFFAIRE'));
          end;
      end;
  finally
    if vQr <> nil then Ferme(vQr);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 19/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGRES.LoadDataMnt;
var
  vSt       : String;
  vQR       : TQuery;
  i         : Integer;
  vTobLigne : Tob;

begin

  vSt := 'SELECT ATA_AFFAIRE, ATA_AFFAIRE0, ATA_AFFAIRE1, ATA_AFFAIRE2,';
  vSt := vSt + ' ATA_AFFAIRE3, ATA_AVENANT, ATA_ARTICLE, GA_LIBELLE, ATA_FONCTION, AFO_LIBELLE,';
  vSt := vSt + ' ATA_NUMEROTACHE, GA_FAMILLENIV1, AFF_LIBELLE, AFF_TIERS,';
  vSt := vSt + ' 0.00 AS PC_MOIS, 0.00 AS PC_MOIS1, 0.00 AS PC_MOIS2, 0.00 AS PC_MOIS3, ';
  vSt := vSt + ' 0.00 AS PC_MOIS4, 0.00 AS PC_MOIS5, 0.00 AS PC_MOIS6, 0.00 AS PC_MOIS7, ';
  vSt := vSt + ' 0.00 AS PC_MOIS8, 0.00 AS PC_MOIS9, 0.00 AS PC_MOIS10, 0.00 AS PC_MOIS11, ';
  vSt := vSt + ' 0.00 AS PC_MOIS12, 0.00 AS PC_MOIS13, 0.00 AS PC_MOIS14, 0.00 AS PC_MOIS15, ';
  vSt := vSt + ' 0.00 AS PC_MOIS16, 0.00 AS PC_MOIS17 ';

  if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
  begin
    vSt := vSt + ' ,0.00 AS PC_MOISCRA, 0.00 AS PC_MOIS1CRA, 0.00 AS PC_MOIS2CRA, ';
    vSt := vSt + ' 0.00 AS PC_MOIS3CRA, 0.00 AS PC_MOIS4CRA, 0.00 AS PC_MOIS5CRA, ';
    vSt := vSt + ' 0.00 AS PC_MOIS6CRA, 0.00 AS PC_MOIS7CRA, 0.00 AS PC_MOIS8CRA, ';
    vSt := vSt + ' 0.00 AS PC_MOIS9CRA, 0.00 AS PC_MOIS10CRA, 0.00 AS PC_MOIS11CRA, ';
    vSt := vSt + ' 0.00 AS PC_MOIS12CRA, 0.00 AS PC_MOIS13CRA, 0.00 AS PC_MOIS14CRA, ';
    vSt := vSt + ' 0.00 AS PC_MOIS15CRA, 0.00 AS PC_MOIS16CRA, 0.00 AS PC_MOIS17CRA ';
  end;

  vSt := vSt + ' FROM AFFAIRE, ARTICLE , TACHE LEFT OUTER JOIN FONCTION ';
  vSt := vSt + ' ON ATA_FONCTION = AFO_FONCTION ';

  vSt := vSt + AddClauseFrom;

  vSt := vSt + ' WHERE ATA_AFFAIRE = AFF_AFFAIRE';
  vSt := vSt + ' AND ATA_TYPEARTICLE <> "PRE" ';
  vSt := vSt + ' AND ATA_ARTICLE = GA_ARTICLE ';

  vSt := vSt + AddClauseWhere('ATA');

  vSt := vSt + ' GROUP BY ATA_AFFAIRE, ATA_AFFAIRE0, ATA_AFFAIRE1,';
  vSt := vSt + ' ATA_AFFAIRE2, ATA_AFFAIRE3, ATA_AVENANT,';
  vSt := vSt + ' ATA_ARTICLE, ATA_FONCTION, AFO_LIBELLE, ';
  vSt := vSt + ' GA_LIBELLE, ATA_NUMEROTACHE, GA_FAMILLENIV1, ';
  vSt := vSt + ' AFF_LIBELLE, AFF_TIERS ';

  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    if Not vQR.Eof then
      begin
         fTobViewer.LoadDetailDB('MONTOBVIEWER','','',vQr,False,True);

        // Ajout dans la liste des données du CRA pour les ressources non planifiées
        if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then AddRessourceCRA(fTobViewer);

        // pour chaque ressource
        for i := 0 to fTobViewer.Detail.Count - 1 do
          begin
            vTobLigne := fTobViewer.Detail[i];
            LoadMois(vTobLigne, vTobLigne, True, vTobLigne.Getvalue('ATA_AFFAIRE'), 'ATA', 'APL_INITPTPR');
 
            if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
              LoadMoisCra(vTobLigne, vTobLigne.Getvalue('ATA_AFFAIRE'));

          end;
        // suppression des lignes
        SuppNonPlanifie(fTobViewer);
      end
    else
      begin
        // Ajout dans la liste des données du CRA pour les ressources non planifiées
        if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then AddRessourceCRA(fTobViewer);

        // pour chaque ressource
        for i := 0 to fTobViewer.Detail.Count - 1 do
          begin
            vTobLigne := fTobViewer.Detail[i];
            LoadMois(vTobLigne, vTobLigne, True, vTobLigne.Getvalue('ATA_AFFAIRE'), 'ATA', 'APL_INITPTPR');

            if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
              LoadMoisCra(vTobLigne, vTobLigne.Getvalue('ATA_AFFAIRE'));

          end;
      end;

  finally
    if vQr <> nil then Ferme(vQr);
  end;
end;

procedure TOF_AFPLANNINGRES.OnArgument (S : String ) ;
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

  TCheckBox(GetControl('CBPDCQUANTITE')).OnClick := CBPDCQUANTITEClick;

end;

procedure TOF_AFPLANNINGRES.CBPDCQUANTITEClick(Sender: TObject);
begin
  if TCheckbox(GetControl('CBPDCQUANTITE')).Checked then
  begin
    SetControlVisible('ATR_RESSOURCE', True);
    SetControlVisible('TATR_RESSOURCE', True);
  end
  else
  begin
    SetControlVisible('ATR_RESSOURCE', false);
    SetControlVisible('TATR_RESSOURCE', false);
    SetControlText('ATR_RESSOURCE', '');
  end;
end;

procedure TOF_AFPLANNINGRES.InitFiltre;
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

procedure TOF_AFPLANNINGRES.OnClose;
begin
  Inherited;
  if assigned(fTobViewer) then
    begin
      fTobViewer.Free;
      fTobViewer := nil;
    end;
end;

procedure TOF_AFPLANNINGRES.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
Begin
  Aff   := THEdit(GetControl('ATA_AFFAIRE'));
  Aff1  := THEdit(GetControl('ATA_AFFAIRE1'));
  Aff2  := THEdit(GetControl('ATA_AFFAIRE2'));
  Aff3  := THEdit(GetControl('ATA_AFFAIRE3'));
  Aff4  := THEdit(GetControl('ATA_AVENANT'));
  Tiers := THEdit(GetControl('ATA_TIERS'));
End;

function TOF_AFPLANNINGRES.AddClauseWhere(pStPrefixe : String) : String;
var
  vStParam : String;

begin
  result := copy(RecupWhereCritere(TPageControl(GetControl('PAGES'))), 6, length(RecupWhereCritere(TPageControl(GetControl('PAGES')))));
  if result <> '' then result := ' AND ' + result;
  if pStPrefixe <> 'ATA' then ReplaceSubStr(result, 'ATA_', pStPrefixe + '_');

  // C.B 27/08/02 ajout des ressources en critere
  // C.B 10/10/02 changement de table suite au remplacement fonction -> article
  ReplaceSubStr(result, 'ATA_RESSOURCE', 'ARS_RESSOURCE');
  ReplaceSubStr(result, 'ATR_RESSOURCE', 'ARS_RESSOURCE');

  vStParam := GetControlText('YTC_TABLELIBRETIERS1');
  if vStParam <> '' then result := result + ' AND YTC_TIERS = AFF_TIERS ';
end;

function TOF_AFPLANNINGRES.AddClauseFrom : String;
var
  vStParam : String;
begin

  vStParam := GetControlText('YTC_TABLELIBRETIERS1');
  if vStParam <> '' then result := result + ', TIERSCOMPL ';

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 25/09/2002
Modifié le ... :   /  /
Description .. : Suppression des lignes n'ayant pas de planification
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGRES.SuppNonPlanifie(pTobLigne : Tob);
var
  i : Integer;
begin
  for i := pTobLigne.detail.count - 1 downto 0 do
    begin
      if (GetParamSoc('SO_AFCLIENT') = cInClientAlgoe) then
      begin
        if  (pTobLigne.detail[i].GetValue('PC_MOIS') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS1') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS2') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS3') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS4') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS5') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS6') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS7') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS8') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS9') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS10') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS11') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOISCRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS1CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS2CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS3CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS4CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS5CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS6CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS7CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS8CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS9CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS10CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS11CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS12CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS13CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS14CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS15CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS16CRA') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS17CRA') = 0) then
         pTobLigne.detail[i].Free;
      end
      else
      begin
        if  (pTobLigne.detail[i].GetValue('PC_MOIS') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS1') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS2') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS3') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS4') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS5') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS6') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS7') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS8') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS9') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS10') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS11') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS12') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS13') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS14') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS15') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS16') = 0) and
            (pTobLigne.detail[i].GetValue('PC_MOIS17') = 0) then
         pTobLigne.detail[i].Free;
      end;
    end;
end;
 
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : on recherche le planifié du CRA et on verifie
                 si la ressource n'est pas déja dans la tob
                 si elle n'y est pas , on l'ajoute

Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGRES.AddRessourceCRA(pTobViewer : Tob);
var
  vSt     : String;
  vTob    : Tob;
  vQr     : TQuery;
  i       : Integer;

begin
                                                                              
  vTob := nil;
  vSt := 'SELECT ACT_AFFAIRE AS ATA_AFFAIRE, ACT_AFFAIRE0 AS ATA_AFFAIRE0, ';
  vSt := vSt + ' ACT_AFFAIRE1 AS ATA_AFFAIRE1, ACT_AFFAIRE2 AS ATA_AFFAIRE2,';
  vSt := vSt + ' ACT_AFFAIRE3 AS ATA_AFFAIRE3, ACT_AVENANT AS ATA_AVENANT, ';
  vSt := vSt + ' ACT_ARTICLE AS ATA_ARTICLE, GA_LIBELLE, ';
  vSt := vSt + ' ACT_FONCTIONRES AS ATA_FONCTION, AFO_LIBELLE,';
  vSt := vSt + ' 0 AS ATA_NUMEROTACHE, ACT_RESSOURCE AS ATA_RESSOURCE, ARS_LIBELLE, ';
  vSt := vSt + ' GA_FAMILLENIV1, AFF_LIBELLE, AFF_TIERS, ';
  vSt := vSt + ' 0.00 AS PC_MOIS, 0.00 AS PC_MOIS1, 0.00 AS PC_MOIS2, 0.00 AS PC_MOIS3, ';
  vSt := vSt + ' 0.00 AS PC_MOIS4, 0.00 AS PC_MOIS5, 0.00 AS PC_MOIS6, 0.00 AS PC_MOIS7, ';
  vSt := vSt + ' 0.00 AS PC_MOIS8, 0.00 AS PC_MOIS9, 0.00 AS PC_MOIS10, 0.00 AS PC_MOIS11, ';
  vSt := vSt + ' 0.00 AS PC_MOISCRA, 0.00 AS PC_MOIS1CRA, 0.00 AS PC_MOIS2CRA, ';
  vSt := vSt + ' 0.00 AS PC_MOIS3CRA, 0.00 AS PC_MOIS4CRA, 0.00 AS PC_MOIS5CRA, ';
  vSt := vSt + ' 0.00 AS PC_MOIS6CRA, 0.00 AS PC_MOIS7CRA, 0.00 AS PC_MOIS8CRA, ';
  vSt := vSt + ' 0.00 AS PC_MOIS9CRA, 0.00 AS PC_MOIS10CRA, 0.00 AS PC_MOIS11CRA, ';
  vSt := vSt + ' 0.00 AS PC_MOIS12CRA, 0.00 AS PC_MOIS13CRA, 0.00 AS PC_MOIS14CRA, ';
  vSt := vSt + ' 0.00 AS PC_MOIS15CRA, 0.00 AS PC_MOIS16CRA, 0.00 AS PC_MOIS17CRA ';
  vSt := vSt + ' FROM AFFAIRE, ACTIVITE, RESSOURCE, FONCTION, ARTICLE ';

  vSt := vSt + AddClauseFrom;

  vSt := vSt + ' WHERE ACT_AFFAIRE = AFF_AFFAIRE';
  vSt := vSt + ' AND ACT_TYPEARTICLE = "PRE" ';
  vSt := vSt + ' AND ACT_ETATVISA = "VIS" ';
  vSt := vSt + ' AND ACT_ARTICLE = GA_ARTICLE ';
  vSt := vSt + ' AND ACT_DATEACTIVITE >= "' + UsDateTime(DebutDeMois(now)) +'"';

  vSt := vSt + AddClauseWhere('ACT');

  vSt := vSt + ' AND ACT_RESSOURCE = ARS_RESSOURCE';
  vSt := vSt + ' AND ACT_FONCTIONRES = AFO_FONCTION';

  vSt := vSt + ' GROUP BY ACT_AFFAIRE, ACT_AFFAIRE0, ACT_AFFAIRE1,';
  vSt := vSt + ' ACT_AFFAIRE2, ACT_AFFAIRE3, ACT_AVENANT,';
  vSt := vSt + ' ACT_ARTICLE, ACT_FONCTIONRES, AFO_LIBELLE, ';
  vSt := vSt + ' ACT_RESSOURCE, ARS_LIBELLE, GA_LIBELLE, GA_FAMILLENIV1,';
  vSt := vSt + ' AFF_LIBELLE, AFF_TIERS ';
   
  vQr := nil;
  Try
    vQr := OpenSql(vSt,True);
    if Not vQr.Eof then
      begin

        vTob :=  TOB.create('QteRealise', nil, -1);
        vTob.LoadDetailDB('MONTOBVIEWER','','',vQr,False,True);
        for i := vTob.Detail.count - 1 downto 0 do
          begin
            if (pTobViewer.FindFirst(['ATA_AFFAIRE', 'ATA_RESSOURCE'], [vTob.Detail[i].GetValue('ATA_AFFAIRE'), vTob.Detail[i].GetValue('ATA_RESSOURCE')], false) = nil) then
              begin
                // ajout de la ligne
                vTob.Detail[i].ChangeParent(pTobViewer, -1);
              end;
          end;
      end;
  Finally
    if vQr <> nil then  Ferme(vQr);
    if vTob <> nil then vTob.Free;
  end;
end;

Initialization
  registerclasses ([TOF_AFPLANNINGRES]) ;
end.



