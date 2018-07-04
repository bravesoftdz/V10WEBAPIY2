{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/04/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : APPRECIATION ()
Mots clefs ... : TOF;APPRECIATION
*****************************************************************}
Unit UTOFAPPRECIATION ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils,  ComCtrls, windows, messages,
{$IFDEF EAGLCLIENT}
   MaineAGL,
{$ELSE}
   dbTables, db,FE_Main,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF ,vierge, HPanel,
     UTOB,Affaireutil,factutil,M3FP,Saisutil,
     paramsoc,DicoAF,AFprepfact,facture,graphics,EntGC,
     AppreciationUtil,factgrp,utilressource,factcomm,UtofAfSaisDate,
     UtofAfPiece,CalcOleGenericAff,UtofAfActiviteMul,UtofAfModifBoni
     ,ConfidentAffaire
     ,UtofAfActiviteMulBm (* pour être sur que la tof soit incluse utiliser dans le script*)
     ;


Type
  TOF_APPRECIATION = Class (TOF)
  	public
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (StArgument : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure CommentaireOnElipsisClick (Sender : TObject);
    function ValiderAppreciation : boolean;

    private
    		Erreur : boolean;
        ExitSansMaj: boolean;
        gCodeAffaire, gDateAppreciation, gDateAppreciationPrec : string;
        gDateDebAffaire , gDateDebExCabinet, gDateDebExCabinet_1  : string;  // Date Debut affaire , début exer ,date appre
        TobAff : TOB;
        TobSituation, TobSitu : TOB;
        TobEchApp, TobEchAppPrec : TOB;
        DEV : RDEVISE;
        BM : R_BM;
        ech_ok ,EcrSimple : boolean;
        TypPrec : string;
        gConsultation : boolean;
        StoDateDeb : TdateTime;
        titre : string;
        Empilage : boolean;
        gbPremiereEcheance : boolean;
        gDateApprecAnneePrec : TDateTime;
        gSortieDirecte : boolean;

    procedure AfterShow;
    procedure FermeTout ;
    procedure OuvreTout ;
    function InfosGene : integer;
    function  RechEcheance(var DatEch : string) : boolean;
    Procedure RecupActiviteTob(TypAct,CodeAff :String ;DateD,DateF:TdateTime;var zz_pq,zz_p,zz_fr,zz_fo : double) ;
    Procedure RecupActivitedansTOB( Req : string;var zqte,zpre,zfra,zfou : Double);

    Procedure RecupBoniMaliTob(SansApprecCourante, consult:boolean;TypAct,CodeAff :String ;DateD,DateF:TdateTime; napp : integer;
				var zz_pq,zz_p,zz_fr,zz_fo,zz2_pq,zz2_p,zz2_fr,zz2_fo: double) ;

    procedure RecupTobLignePiece(AvecFactApprec, consult:boolean;CodeAff,znat,typgene : String ;zrep:Integer;DateD,DateF:TdateTime; var zz_pq,zz_p,zz_fr,zz_fo : double) ;
    Function  FabriqueReqPiece(typ : string;consult : boolean) : string;

    procedure InfosSituationPeriode;
    Procedure CreatTobSituation;
    Procedure InitTobSituation(nomdeb,nomfin : string);
    Procedure AfficheEcranSituation;
    Procedure GestionZoneEcranBM(Num : Integer);
    Procedure Cumulchamp(nom : string);
    Procedure AlimChampTobSituation(nom :string;zz,zz_pq,zz_p,zz_fr,zz_fo : double);
    Procedure CumECBRU;
    Procedure CumPRNET;
    Procedure CumANPRNET;
    Procedure CumANECPRNET;
    Procedure CumECNET;
    Procedure CumANECNET;
    Procedure CumANECECNET;


// MAJ
    procedure AcceslignesVisa(num : integer);
    procedure AppelProd(typ : string);
    procedure AppelFact(typ : string; periode : string);
    procedure ModifMontAFacturer;
    procedure RepriseEncours;
//    procedure RepriseANouveau;
    procedure Actualisation;
		Procedure CalcEC(zecnet,zafac,zbon1,zbon2,zec : string);

    procedure InfosDecision(num :integer);
    Procedure AfficheMontDecision;
    Procedure ControleSaisie;
    Procedure ModifZoneDeci(zone : string;ind : integer;zenable ,zinit: boolean;zcolor: Tcolor);
    Procedure InitZoneDeci(zone : string;ind : integer);


    Procedure RecupSaisiedansFactaff;
    procedure ValideVisa;
    Procedure PrepaFacpro;
    Procedure TrtTypeFacture;
		Function  ChangementTypeFact:boolean;
    procedure TrtSpecifSolde ;
    Procedure SoldeCreatBoniMali;
    Procedure RecupMontantLignesvisa;
    Procedure AppelAppreciationPrecedente;
    Procedure AppelAppreciationAnneePrecedente;
    Procedure AppelModifBoniMali;
    Procedure CalculeAFacturer;


    Procedure SupprApprec;
    procedure RAZApprec ;
    function MajFactaff : boolean;
    function GenerFacture : boolean;
    procedure ValideActivitePiece ;
    procedure AnnulMaj ;
    Procedure MajAffaireDate;

    Procedure FormCloseQuery (Sender : TObject; var CanClose : Boolean);
  end ;

  const
  Suff : array[1..5] of string  = ('','_PQ','_P','_FR','_FO') ;
  MaxSuff : integer = 5;

    	// libellés des messages de la TOF  AfFactDecideur
	TexteMsgAffaire: array[1..10] of string 	= (
          {1}        'La saisie des zones "à facturer" est incohérente'
          {2}        ,'Appréciation visée donc non modifiable'
          {3}        ,'La saisie des montants  est incohérente'
          {4}        ,'Vous ne pouvez pas viser une appréciation sans donnée '
          {5}        ,'Confirmez-vous le changement de mode de facturation ?'
          {6}        ,'Génération en cours'
          {7}        ,'Vérifiez vos dates de génération'
          {8}        ,'Vous devez sélectionner un mode de visa avant de valider'
          {9}        ,'Il n''existe pas d''appréciation antérieure.'
          {10}       ,'Il n''existe pas d''appréciation le même mois de l''année dernière.'
                     );

Function AFLanceFiche_AppreciationNiv0 (Argument:string):variant;
Function AFLanceFiche_AppreciationCompl (Argument:string):variant;
Procedure AnnulationApprec (affaire : string;NumEche:integer; Simulation : boolean);
Procedure MajPieceAvecAncienNumApprec(affaire : string;NumEche:integer);


Implementation



procedure TOF_APPRECIATION.OnLoad;
begin
  Inherited;

  titre := ecran.caption;
  Erreur := false;
  ExitSansMaj := true;
  if Empilage then
    begin
      Ecran.Top := Ecran.Top + 20;
      Ecran.Left := Ecran.Left + 20;
    end;

  // init des infos bulles
  SetControlProperty ('AFA_JUSTIFBONI', 'Hint', RechDom('AFJUSTIFBONI', GetControlText ('AFA_JUSTIFBONI'), false));
  SetControlProperty ('AFA_JUSTIFBONI2', 'Hint', RechDom('AFJUSTIFBONI', GetControlText ('AFA_JUSTIFBONI2'), false));
end ;

procedure TOF_APPRECIATION.OnArgument (StArgument : String ) ;
  Var
    Critere, Champ, valeur  : String;
    repInfoGene, x : integer;
    zori : string;    // Origine de l'appel: appreciation simple ou complete
begin
  Inherited ;
  gSortieDirecte := false;
  gConsultation := false;
  Empilage := false;
  gbPremiereEcheance := false;
  gDateApprecAnneePrec := idate1900;
  gDateAppreciationPrec := Datetostr(idate1900);
  If (GetControl ('DATEAPPRECPREC') <> Nil) then
    SetControlText ('DATEAPPRECPREC', gDateAppreciationPrec);

    // Recup des critères
  Critere:=(Trim(ReadTokenSt(stArgument)));
  While (Critere <>'') do
      BEGIN
      if Critere<>'' then
          BEGIN
          X:=pos(':',Critere);
          if x<>0 then
             begin
             Champ:=copy(Critere,1,X-1);
             Valeur:=Copy (Critere,X+1,length(Critere)-X);
             end;
          if Champ = 'ZORI' then
            Zori := Valeur;
          if Champ = 'ZAFF' then
            gCodeAffaire := Valeur;
          if Champ = 'ZDATAPP' then
            gDateAppreciation := Valeur;
          if Champ = 'CONSULT' then
            gConsultation := true;

          END;
      Critere:=(Trim(ReadTokenSt(stArgument)));
      END;

      TypPrec := '';
      // Ecran simple ou ecran Mazard
      if (copy(zori, 1, 3) = 'SIM') then EcrSimple := true else EcrSimple := false;

      // Si on empile les ecrans d'appreciation (pour consulter les precedentes)
      // Certaines zones et certains champs sont inaccessibles
      if (length(zori) > 3) then
        begin
          Empilage := true;
          SetControlEnabled ('BSTOP', false);
          SetControlEnabled ('BPOPMENU', false);
          SetControlProperty('BENC', 'Enabled', false);
          SetControlProperty('BEFFACE', 'Enabled', false);
          SetControlProperty('BAFAC', 'Enabled', false);
          if (GetControlText('ZABON1') <> '') or (GetControlText('ZABON2') <> '') then
            SetControlProperty('BMODBONI', 'Enabled', true);
        end;

      TFVierge(Ecran).OnAfterFormShow := AfterShow;

      // Creation des TOB
      OuvreTout;

      // Initialisations générales (structure Boni/mali, Tob echeances, dates et libellés)
      repInfoGene := InfosGene;
      if (repInfoGene = 1) then
        // L'appréciation précédente n'a pas été visée, on ouvre celle-ci pour la viser
        begin
          gCodeAffaire := TOBEchAppPrec.GetValue('AFA_AFFAIRE');
          gDateAppreciation := DateToStr(TOBEchAppPrec.GetValue('AFA_DATEECHE'));
          InfosGene;
        end
      else
      if (repInfoGene = 2) then
        // L'appréciation précédente n'a pas été visée, on ne veut pas continuer
        begin
          gSortieDirecte := true;
          exit;
        end;

      // Cas de consultation
      if ((TobEchApp.GetValue('AFA_ETATVISA') = 'VIS') or Empilage or gConsultation) then
        gConsultation := true
      else
        gConsultation := false;

      // Recherche des infos de la zone Situation (Période dans le cas du Solde)
      InfosSituationPeriode;

      // Recherche des infos de la zone Décision
      InfosDecision (0);

      //RepriseANouveau;

      // zone saisissable
      SetFocusControl('ZAFAC_P');
      nextControl(Ecran);
      SetFocusControl('ZAFAC');

      // Gestion des commentaires
      if GetParamSoc('SO_GCCOMMENTAIRE') then
        begin
          SetControlProperty('AFA_LIBELLEMULTI1', 'Width', THEdit(GetControl('AFA_LIBELLE1')).Width);
          SetControlProperty('AFA_LIBELLEMULTI2', 'Width', THEdit(GetControl('AFA_LIBELLE2')).Width);
          SetControlProperty('AFA_LIBELLEMULTI3', 'Width', THEdit(GetControl('AFA_LIBELLE3')).Width);
          SetControlVisible('AFA_LIBELLE1', false);
          SetControlVisible('AFA_LIBELLE2', false);
          SetControlVisible('AFA_LIBELLE3', false);
          SetControlVisible('AFA_LIBELLEMULTI1', true);
          SetControlVisible('AFA_LIBELLEMULTI2', true);
          SetControlVisible('AFA_LIBELLEMULTI3', true);
          THEdit(GetControl('AFA_LIBELLEMULTI1')).OnElipsisClick := CommentaireOnElipsisClick;
          THEdit(GetControl('AFA_LIBELLEMULTI2')).OnElipsisClick := CommentaireOnElipsisClick;
          THEdit(GetControl('AFA_LIBELLEMULTI3')).OnElipsisClick := CommentaireOnElipsisClick;
        end;

      if (gConsultation) then
      Begin
        CalcEC('ZECNET','ZAFAC','ZABON1','ZABON2','ZEC');

        SetControlProperty('BENC', 'Enabled', false);
        SetControlProperty('BEFFACE', 'Enabled', false);
        SetControlProperty('BAFAC', 'Enabled', false);
        SetControlProperty('BSTOP', 'visible', false);
        SetControlProperty('BCONSBONI', 'Enabled', true);

        SetControlProperty('TZAFAC','Color',clbtnface);
        if THEdit(GetControl('ZRACPTE')) <> nil then
         begin  // n'existe que sur fiche AfAppreComplete
         SetControlProperty('TZRACPTE','Color',clbtnface);
         ModifZoneDeci('ZRACPTE',1,false,false,clbtnface);
         end;

        SetControlProperty('TZMOINS','Color',clbtnface);
        ModifZoneDeci('ZAFAC',1,false,false,clbtnface);
        SetControlEnabled('AFA_BLOCNOTE',false);SetControlProperty('AFA_BLOCNOTE','Color',clbtnface);
        SetControlEnabled('AFA_LIBELLE1',false);SetControlProperty('AFA_LIBELLE1','Color',clbtnface);
        SetControlEnabled('AFA_LIBELLE2',false);SetControlProperty('AFA_LIBELLE2','Color',clbtnface);
        SetControlEnabled('AFA_LIBELLE3',false);SetControlProperty('AFA_LIBELLE3','Color',clbtnface);
        SetControlEnabled('AFA_LIBELLEMULTI1',false);SetControlProperty('AFA_LIBELLEMULTI1','Color',clbtnface);
        SetControlEnabled('AFA_LIBELLEMULTI2',false);SetControlProperty('AFA_LIBELLEMULTI2','Color',clbtnface);
        SetControlEnabled('AFA_LIBELLEMULTI3',false);SetControlProperty('AFA_LIBELLEMULTI3','Color',clbtnface);
        SetControlEnabled('AFA_PROFILGENER',false);SetControlProperty('AFA_PROFILGENER','Color',clbtnface);
        SetControlEnabled('AFA_ETATVISA',false);SetControlProperty('AFA_ETATVISA','Color',clbtnface);

        ModifZoneDeci('ZABON1',1,false,false,clbtnface);
        ModifZoneDeci('ZABON2',1,false,false,clbtnface);

        SetControlEnabled('ZAFAC_PQ1',false);SetControlProperty('ZAFAC_PQ1','Color',clbtnface);
        SetControlEnabled('ZABON1_PQ1',false);SetControlProperty('ZABON1_PQ1','Color',clbtnface);
        SetControlEnabled('ZABON2_PQ1',false);SetControlProperty('ZABON2_PQ1','Color',clbtnface);
        SetControlEnabled('AFA_JUSTIFBONI',false);SetControlProperty('AFA_JUSTIFBONI','Color',clbtnface);
        SetControlEnabled('AFA_JUSTIFBONI2',false);SetControlProperty('AFA_JUSTIFBONI2','Color',clbtnface);

        SetControlEnabled('BVALIDER',false);
        SetControlEnabled('BDELETE',false);
        SetControlEnabled('BLETTR',false);
        if (GetControlText('ZABON1') <> '') or (GetControlText('ZABON2') <> '') then
          SetControlProperty('BMODBONI', 'Enabled', true);


      if not (EcrSimple) then
        Begin
          SetControlEnabled('CHACPT', false);
          SetControlEnabled('CHFAC', false);
          SetControlEnabled('CHSOLD', false);
        End;

      End;

  TForm(Ecran).OnCloseQuery := FormCloseQuery;
end ;



procedure TOF_APPRECIATION.AfterShow;
begin

if gSortieDirecte then
  begin
    if IsInside (TFVierge (Ecran)) then
      begin
        THPanel (Ecran.Parent).CloseInside;
        THPanel (Ecran.Parent).VideToolBar;
      end;

    PostMessage(TFVierge(Ecran).Handle, WM_CLOSE, 0, 0);
  end;

end;


procedure TOF_APPRECIATION.CommentaireOnElipsisClick (Sender : TObject);
var
  CodeCom : string;
  QQ : TQuery;
begin
  CodeCom := AGLLanceFiche ('GC', 'GCMULCOMMENTAIRE', CodeCom, '', '');

  if CodeCom <> '' then
    begin
      try
        QQ := OpenSql('Select GCT_LIBELLE From COMMENTAIRE WHERE GCT_CODE="' + CodeCom + '"', True);
        if not QQ.Eof then
          THEdit(Sender).text := QQ.FindField('GCT_LIBELLE').AsString;
      finally
        Ferme(QQ);
      end;
    end;
end;

// Appuie sur le bouton de validation
procedure TOF_APPRECIATION.OnUpdate;
begin
  Inherited;

  Erreur := false;

  //controles
	ControleSaisie;
  if Erreur then exit;

  V_PGI.IoError := oeOK;
  // Validation
  try
    BeginTrans;
    if (ValiderAppreciation = true) then
      CommitTrans
    else
      V_PGI.IoError := oeUnknown;

  except
      V_PGI.IoError := oeUnknown;
  end;

  if (V_PGI.IoError = oeUnknown) then
      begin
        RollBack;
        exit;
      end;

  // Sortie
  ExitSansMaj := false;
  PGIInfoAF ('Appréciation bien enregistrée', titre);
  Ecran.Close;
END;

function TOF_APPRECIATION.ValiderAppreciation : boolean;
begin
  Result := true;

  // Mises à jour liées au factaff
	if not MajFactaff then
    begin
      Result := false;
      exit;
    end;

  // création des factures Pro puis définitives
  if not GenerFacture then
    Result := false;

end;

procedure TOF_APPRECIATION.OnDelete ;
begin
  Inherited ;
//  	if (GetControlText('AFA_ETATVISA') = 'VIS') then
  If (PGIAskAF ('Confirmez-vous la suppression de l''appréciation en cours ?', titre) = mrYes) then
    Begin
      SupprApprec;
      Erreur := false;
      ExitSansMaj := false;
    End;
  Ecran.Close;
end ;

procedure TOF_APPRECIATION.FormCloseQuery (Sender : TObject; var CanClose : Boolean);
begin

  if (ExitSansMaj) and not gConsultation then
    Begin
      If (PGIAskAF ('Confirmez-vous l''annulation des modifications en cours ?', titre) = mrYes) then
        // Gestion d'une sortie simple sans maj, defaire les éventuels lettrages déjà faits
        begin
          V_PGI.IoError := oeOK;
          try
            BeginTrans;

            AnnulMaj;

            if (V_PGI.IoError = oeOK) then
              CommitTrans;

          except
              V_PGI.IoError := oeUnknown;
          end;

          if (V_PGI.IoError = oeUnknown) then
              begin
                RollBack;
                PGIInfoAF ('Les modifications en cours n''ont pu être annulées.', titre);
              end;

          CanClose := true;
        end
      else
        CanClose := false;
    End;

end;

procedure TOF_APPRECIATION.OnClose;
begin
  FermeTout;

  Inherited;
end ;

procedure TOF_APPRECIATION.OuvreTout ;
BEGIN
  TOBAff:=TOB.Create('AFFAIRE',Nil,-1);

  TobSituation:=TOB.Create('Zones Situation',Nil,-1);
  TobSitu := TOB.create ('les zones',TobSituation,-1);
  CreatTobSituation;

  TobEchApp :=TOB.Create('FACTAFF',Nil,-1);
  TobEchAppPrec :=TOB.Create('FACTAFF',Nil,-1);
END;

procedure TOF_APPRECIATION.FermeTout ;
BEGIN
  TobAff.free;  TobAff := NIL;
  TobSituation.free;  TobSituation := NIL;

  TobEchApp.free; TobEchApp := NIL;
  TobEchAppPrec.free; TobEchAppPrec := NIL;
END;

//******************************************************************
//          MAJ , ANNULATION  ETC....
//******************************************************************


{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 22/10/2001
Modifié le ... :   /  /
Description .. : Suppression d'une appréciation ,
Suite ........ : raz num apprec dans activite,piece, suppr factaff
Suite ........ : et boni/mali
Mots clefs ... : GIGA;APPRECIATION;SUPPR
*****************************************************************}
procedure TOF_APPRECIATION.SupprApprec ;
BEGIN
  try
    BeginTrans;

    AnnulationApprec (gCodeAffaire, TobEchApp.GetValue('AFA_NUMECHE'), EcrSimple);

    if (V_PGI.IoError = oeOK) then
      CommitTrans
    else
      begin
        RollBack;
        exit;
      end;

  except
     RollBack;
  end;
END;

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 16/10/2001
Modifié le ... :   /  /    
Description .. : Si modif Type de facture, raz des lignes dèja topées...
Mots clefs ... : GIGA;APPRECIATION
*****************************************************************}
procedure TOF_APPRECIATION.RAZApprec ;
var Req : string;
		napp : integer;
BEGIN
  napp := TobEchApp.GetValue('AFA_NUMECHE');
  req := 'UPDATE ACTIVITE  SET act_etatvisafac="ATT",act_numapprec=0,act_activiterepris="F" '
  +' WHERE act_affaire="'+gCodeAffaire+'"'+ ' AND ((act_numapprec="'+ inttostr(napp)+'")'
  + ' OR (act_numapprec="'+ inttostr(OffSetApp+napp)+'"))';
  ExecuteSQL(req);

    // recherche ancien num apprec
  MajPieceAvecAncienNumApprec(gCodeAffaire,napp);

  InitZoneDeci('ZAFAC',1);
  InitZoneDeci('ZABON1',1);
  InitZoneDeci('ZABON2',1);
  SetControlText('AFA_JUSTIFBONI','');
  SetControlText('AFA_JUSTIFBONI2','');

END;

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 15/10/2001
Modifié le ... :   /  /    
Description .. : Annulation des visas dèjà fait en cours de l'appréciation, en 
Suite ........ : cas de sortie simple  sans maj
Mots clefs ... : GIGA;APPRECIATION
*****************************************************************}
procedure TOF_APPRECIATION.AnnulMaj;
  var Req, typgen : string;
BEGIN
  Typgen := TobEchApp.GetValue('AFA_GENERAUTO');

  if (Typgen  = 'SOL') or (typgen = 'ACT') then
    begin
      req := 'UPDATE ACTIVITE  SET act_etatvisafac="ATT",act_numapprec=0,act_activiterepris="F"'
      +' WHERE act_affaire="' + gCodeAffaire + '"' + ' AND act_numapprec="' + inttostr (OffSetApp + TobEchApp.GetValue('AFA_NUMECHE')) + '"';
      ExecuteSQL (req);

      // voir ce que je fais pour les acomptes
      req := 'UPDATE PIECE  SET gp_facreprise = (-1) '
      + ' WHERE gp_affaire="' + gCodeAffaire + '"'
      + ' AND  gp_naturepieceg ="FAC" AND gp_facreprise="' + IntToStr (OffSetApp + TobEchApp.GetValue('AFA_NUMECHE')) + '"';
      ExecuteSQL (req);
    end;

  if (gbPremiereEcheance = true) then
    begin
      req := 'DELETE FACTAFF '
          + ' WHERE AFA_TYPECHE="APP" AND AFA_AFFAIRE="' + gCodeAffaire + '"' + ' AND AFA_NUMECHE="' + inttostr (TobEchApp.GetValue('AFA_NUMECHE')) + '"';
      ExecuteSQL (req);
    end;
END;

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 23/10/2001
Modifié le ... : 30/07/2003
Description .. : Maj des facture d'acompte, avec leur numéro d'échéance 
Suite ........ : d'origine
Suite ........ : Mis ede cette fct en public ....
Mots clefs ... : GIGA;APPRECIATION;MAJPIECE
*****************************************************************}
Procedure  MajPieceAvecAncienNumApprec (Affaire:string; Numeche:Integer) ;
var Req ,numdoc: string;
		TobPiece,TobDet : Tob;
    QQ : Tquery;
    wi,oldapp : Integer;
BEGIN
	TobPiece := TOB.Create('liste des pieces',Nil,-1);
	Req := 'SELECT GP_NATUREPIECEG,GP_DATEPIECE,GP_SOUCHE,GP_NUMERO,GP_INDICEG,GP_FACREPRISE from PIECE '
  + ' where GP_AFFAIRE = "'+Affaire+'"'+' and (GP_FACREPRISE ="' + inttostr(NumEche)
  + '" or GP_FACREPRISE ="'+inttostr(OffSetApp+NumEche)+'")';
  QQ := nil;
	Try
    QQ := OpenSQL(Req,true);
    If Not QQ.EOF then TobPiece.LoadDetailDB('PIECE','','',QQ,True);
  Finally
    Ferme(QQ);
  End;

  For wi:=0  to  TobPiece.detail.count-1 do
  Begin
  	OldApp := (-1);
  	TobDet := TobPiece.detail[wi];
  	Numdoc := EncodeRefPiece(TobDet);
    Req := 'SELECT AFA_NUMECHE from FACTAFF where AFA_NUMPIECE ="'+Numdoc+'"'
    + 'and AFA_AFFAIRE = "'+Affaire+'"';  // Pas triés ?????
    QQ := nil;
    try
    // PL le 09/01/02 pour optimiser le nombre d'enregistrments lus en eagl
    QQ:=OpenSQL(Req,True,1);
//    QQ:=OpenSQL(Req,True);

  	if Not QQ.EOF then
       begin
       QQ.First;  // classées par ordre décroissant
       OldApp:= QQ.Fields[0].AsInteger;
       end;
    finally
    Ferme(QQ);
    end;

    TobDet.PutValue('GP_FACREPRISE',OldApp*(-1));

    end;
    TobPiece.UpdateDB;
    TobPiece := NIL; TobPiece.free;
END;


{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 12/10/2001
Modifié le ... : 12/10/2001
Description .. : Mise à jour de l'activité et de la facturation.
Suite ........ : Si validation maj du n° echéance
Mots clefs ... : GIGA;APPRECIATION;MAJ
*****************************************************************}
procedure TOF_APPRECIATION.ValideActivitePiece ;
var
  Req : string;
BEGIN

  Req := 'UPDATE ACTIVITE SET act_activiterepris="FAC",'
  +'act_numapprec="'+ inttostr(TobEchApp.GetValue('AFA_NUMECHE'))+'"'
  +' WHERE act_affaire="'+ gCodeAffaire + '"' + ' AND act_numapprec="'+ inttostr(OffSetApp+TobEchApp.GetValue('AFA_NUMECHE'))+'"';
  ExecuteSQL (req);

  Req := 'UPDATE PIECE  SET gp_facreprise ="'+IntToStr(TobEchApp.GetValue('AFA_NUMECHE'))+ '"'
  + ' WHERE gp_affaire="'+ gCodeAffaire + '"'
  + ' AND  gp_naturepieceg ="FAC" AND gp_facreprise="'+IntToStr(OffSetApp+TobEchApp.GetValue('AFA_NUMECHE'))+ '"';
  ExecuteSQL (req);

END;

function TOF_APPRECIATION.MajFactaff : boolean;
BEGIN
  Result := true;

  try
  // mises à jour  activite et piece si besoin
	ValideActivitePiece;

  // maj factaff
  if V_PGI.IoError = oeOk then
    begin
      RecupSaisieDansFactaff;
      ValideVisa;
      TOBEchApp.SetAllModifie(True);

      if Not TOBEchApp.InsertOrUpdateDB (False) then
        Result := false;
    end;

  except
    Result := false;
  end;

END;


Procedure TOF_APPRECIATION.RecupSaisiedansFactaff;
Begin

// Recup des montants depuis le Tob Decision
  TobEchApp.PutValue('AFA_AFACTOT',valeur(GetControlText('ZAFAC')));
	TobEchApp.PutValue('AFA_AFACTURERQTE',0);
//PQ  TobEchApp.PutValue('AFA_AFACTURERQTE',valeur(GetControlText('ZAFAC_PQ')));
  TobEchApp.PutValue('AFA_AFACTURER',valeur(GetControlText('ZAFAC_P')));
  TobEchApp.PutValue('AFA_AFACTFO',valeur(GetControlText('ZAFAC_FO')));
  TobEchApp.PutValue('AFA_AFACTFR',valeur(GetControlText('ZAFAC_FR')));

  if not(EcrSimple) and  TCheckBox(GetControl('CHSOLD')).Checked  then
  Begin
  	TobEchApp.PutValue('AFA_ACPTEQTE',0);
//PQ    TobEchApp.PutValue('AFA_ACPTEQTE',valeur(GetControlText('ZRACPTE_PQ')));
    if THEdit(GetControl('ZRACPTE')) <> nil then
     begin  // n'existe que sur fiche AfAppreComplete
     TobEchApp.PutValue('AFA_ACPTEPR',valeur(GetControlText('ZRACPTE_P')));
     TobEchApp.PutValue('AFA_ACPTEFO',valeur(GetControlText('ZRACPTE_FO')));
     TobEchApp.PutValue('AFA_ACPTEFR',valeur(GetControlText('ZRACPTE_FR')));
     end;
   ENd;

  // Conversion des montants en EURO et DEVISE
  ConvertChampTob('AFA_AFACTOT',DEV,TobEchapp);
  ConvertChampTob('AFA_AFACTURER',DEV,TobEchapp);
  ConvertChampTob('AFA_AFACTFR',DEV,TobEchapp);
  ConvertChampTob('AFA_AFACTFO',DEV,TobEchapp);
  ConvertChampTob('AFA_AFACTURER',DEV,TobEchapp);

  ConvertChampTob ('AFA_ACPTEPR', DEV, TobEchapp);

  if GetParamSoc ('SO_GCCOMMENTAIRE') then
    begin
      TobEchApp.PutValue ('AFA_LIBELLE1', GetControlText('AFA_LIBELLEMULTI1'));
      TobEchApp.PutValue ('AFA_LIBELLE2', GetControlText('AFA_LIBELLEMULTI2'));
      TobEchApp.PutValue ('AFA_LIBELLE3', GetControlText('AFA_LIBELLEMULTI3'));
    end
  else
    begin
      TobEchApp.PutValue ('AFA_LIBELLE1', GetControlText('AFA_LIBELLE1'));
      TobEchApp.PutValue ('AFA_LIBELLE2', GetControlText('AFA_LIBELLE2'));
      TobEchApp.PutValue ('AFA_LIBELLE3', GetControlText('AFA_LIBELLE3'));
    end;
  TobEchApp.PutValue('AFA_BLOCNOTE',GetControlText('AFA_BLOCNOTE'));
  TobEchApp.PutValue('AFA_PROFILGENER',GetControlText('AFA_PROFILGENER'));

  if BM.bm1 then
    begin
    TobEchApp.PutValue('AFA_JUSTIFBONI',GetControlText('AFA_JUSTIFBONI'));
    TobEchApp.PutValue('AFA_BONIMALIQTE',0);
//PQ    TobEchApp.PutValue('AFA_BONIMALIQTE',valeur(GetControlText('ZABON1_PQ')));
    TobEchApp.PutValue('AFA_BONIMALI',valeur(GetControlText('ZABON1_P')));
    TobEchApp.PutValue('AFA_BM1FO',valeur(GetControlText('ZABON1_FO')));
    TobEchApp.PutValue('AFA_BM1FR',valeur(GetControlText('ZABON1_FR')));
  	ConvertChampTob('AFA_BONIMALI',DEV,TobEchapp);
  	ConvertChampTob('AFA_BM1FR',DEV,TobEchapp);
  	ConvertChampTob('AFA_BM1FO',DEV,TobEchapp);
    End;

  if BM.bm2 then
    begin
    TobEchApp.PutValue('AFA_JUSTIFBONI2',GetControlText('AFA_JUSTIFBONI2'));
    TobEchApp.PutValue('AFA_BM2PRQTE',0);
//PQ    TobEchApp.PutValue('AFA_BM2PRQTE',valeur(GetControlText('ZABON2_PQ')));
    TobEchApp.PutValue('AFA_BM2PR',valeur(GetControlText('ZABON2_P')));
    TobEchApp.PutValue('AFA_BM2FO',valeur(GetControlText('ZABON2_FO')));
    TobEchApp.PutValue('AFA_BM2FR',valeur(GetControlText('ZABON2_FR')));

  	ConvertChampTob('AFA_BM2PR',DEV,TobEchapp);
  	ConvertChampTob('AFA_BM2FR',DEV,TobEchapp);
  	ConvertChampTob('AFA_BM2FO',DEV,TobEchapp);
    End;
End;

// Validation visa appréciation
procedure TOF_APPRECIATION.ValideVisa;
Begin
   if GetControlText ('AFA_ETATVISA') = 'VIS' then
  Begin
    TobEchApp.PutValue ('AFA_ETATVISA', GetControlText ('AFA_ETATVISA'));
    TobEchApp.PutValue ('AFA_VISEUR', V_PGI.User);
    TobEchApp.PutValue ('AFA_DATEVISA', date);
  End
  else
  Begin
   	TobEchApp.PutValue ('AFA_ETATVISA', GetControlText ('AFA_ETATVISA'));
    TobEchApp.PutValue ('AFA_VISEUR', '');
//    TobEchApp.PutValue('AFA_DATEVISA',datetostr(idate1900));
  End;

End;


function TOF_APPRECIATION.GenerFacture : boolean;
var
  rep : word;
BEGIN
  Result := true;
  try
    if (TobEchApp.GetValue ('AFA_ETATVISA') = 'VIS') then
      begin
        if (valeur (GetControlText ('ZAFAC')) <> 0) then
          begin
            rep := HShowMessage ('0;' + titre + ';Vous venez de viser l''appréciation, souhaitez-vous générer la facture provisoire ?;Q;YNC;N;N;', '', '');

            if (rep = mrCancel) then
              begin
                Result := false;
              end
            else
            if (rep = mrYes) then
              begin
                PrepaFacpro;
              end;
          end
        else
          begin
            GenerationBoniMali (TobEchApp, tobaff, '');
            // PL le 19/08/03 : dans le cas ou l'on vise mais ou l'on ne génère pas de facture
            // on met quand même à jour la date de dernière appréciation et le numéro d'appréciation dans la mission
            MajAffaireDate;
          end;
      end;
  except
    Result := false;
  end;
END;

Procedure  TOF_APPRECIATION.MajAffaireDate;
var
  Req : string;
begin

  Req := 'UPDATE AFFAIRE SET AFF_DATESITUATION ="' + UsDateTime (TobEchApp.GetValue ('AFA_DATEECHE')) + '"'
     + ',AFF_NUMSITUATION ="' + inttostr(TobEchApp.GetValue('AFA_NUMECHE')) + '"'
     + ' WHERE AFF_AFFAIRE=' + '"' + TobEchApp.GetValue('AFA_AFFAIRE') + '"';


  ExecuteSQL (Req);
end;



Procedure  TOF_APPRECIATION.SoldeCreatBoniMali;
Var zone, zi :double;
		ind : char;
BEGIN
  CalcEC('ZECNET','ZAFAC','ZABON1','ZABON2','ZEC');
  zone:= valeur(GetControlText('ZEC'));
  if zone = 0 then exit;

  if (zone > 0) then
  begin
  	zi := 1;
  	if  (BM.sens1 = (-1)) then ind := '1' else   ind := '2';
  End
  else
  begin
    zi := (-1);
  	if ( BM.sens1 = (1) ) then ind := '1' else ind := '2';
  End;

  SetControlText('ZABON'+ind,strf00(valeur(GetControlText('ZEC'))*zi ,DEV.decimale) );
//PQ  SetControlText('ZABON'+ind+'_PQ',strf00(valeur(GetControlText('ZEC_PQ'))*zi,DEV.decimale));
  SetControlText('ZABON'+ind+'_P',strf00(valeur(GetControlText('ZEC_P'))*zi,DEV.decimale));
  SetControlText('ZABON'+ind+'_FO',strf00(valeur(GetControlText('ZEC_FO'))*zi,DEV.decimale));
  SetControlText('ZABON'+ind+'_FR',strf00(valeur(GetControlText('ZEC_FR'))*zi,DEV.decimale));

  CalcEC('ZECNET','ZAFAC','ZABON1','ZABON2','ZEC');

  END;


Procedure  TOF_APPRECIATION.PrepaFacpro;
var TobRegAff,TobFRegAff,TobProfil,TobArticles : TOB;
		TobComplete,TobPiece : TOB;
    GenFac : R_GENFAC;
    nb_aff,nb_fac,fac_deb,fac_fin : Integer;
    stdate,sttexte,numpiece,zecle,zdate,st: string;
    AvecTrans ,ret: boolean;
    xnat,xsouche,xnat2,xlib,xlib2,xlib3 : string;
    tmp : array of variant;
BEGIN

	TOBArticles :=TOB.Create('ARTICLES',Nil,-1) ;
  TobRegAff   := Tob.Create('Affaire Select',NIL,-1);
  TobFregaff  := TOB.create ('les zones',TobRegAff,-1);
  TobProfil   := Tob.Create('les profils',NIL,-1);

  try
    //chargement de tous les profils, afin de ne pas avoir à les relire a chaque affaire
    // on passera la tobprofil en paramètre

    ret := ChargeTobProfil(TobProfil);

   // xnat := 'FPR'; xsouche := 'GFP';
    if (valeur(GetControlText('ZAFAC')) > 0) then
    begin
      xnat := 'FPR'; xsouche := 'GFP'; xnat2 := 'FAC';
      xlib := 'Facture';
      xlib2 :='la facture définitive';
      xlib3 := 'la facture';
    end
    else
    begin
      xnat := 'APR'; xsouche := 'APR'; xnat2 := 'AVC';
      xlib := 'Avoir';
      xlib2 :='l''avoir définitif';
      xlib3 := 'l''avoir';
    end;

    if not(ret) then
    begin
      StTexte := 'Il y a un code article inconnu dans vos profils';
      PGIInfoAf(StTexte,Titre);
    End
    else
    Begin
       // appel préparation facture Appréciation
      TobFRegAff.AddChampSup('AFF_AFFAIRE',False);
      TobFRegAff.AddChampSup('AFF_PROFILGENER',False);
      TobFRegAff.AddChampSup('T_TIERS',False);
      // mcd 20/06/02 TobFRegAff.AddChampSup('AFF_GENERAUTO',False);
      TobFRegAff.AddChampSup('AFA_GENERAUTO',False);
      TobFregAff.PutValue( 'AFA_GENERAUTO',''); //mcd 22/10/02 ilf aut initialiser la zone qui est utiliser dans ChargeEcheance
      TobFregAff.PutValue( 'AFF_AFFAIRE',TobAff.GetValue('AFF_AFFAIRE'));
      TobFRegAff.PutValue('AFF_PROFILGENER',TobEchApp.GetValue('AFA_PROFILGENER'));
      TobFRegAff.PutValue('T_TIERS',TobAff.GetValue('AFF_TIERS'));
      TobFRegAff.AddChampSup('AFF_PRINCIPALE',False);
      TobFRegAff.PutValue('AFF_PRINCIPALE',TobAff.GetValue('AFF_PRINCIPALE'));  //mcd 25/11/02 pour être en phase avec prépa auto ou champ oblig
      TobFRegAff.AddChampSup('AFF_REGSURCAF',False);
      TobFRegAff.PutValue('AFF_REGSURCAF',TobAff.GetValue('AFF_REGSURCAF'));    //mcd 25/11/02 pour être en phase avec prépa auto ou champ oblig
      Genfac.zgen :=  'G';
      Genfac.ztypech :=  'APP';
      Genfac.znat :=  xnat;
      nb_aff := 1;
      stdate := gDateAppreciation;
      AvecTrans := false;
      AFPrepFact_init(TobRegAff,TobProfil,Nil,Tobarticles,GenFac,nb_aff,stdate,avectrans,nb_fac,fac_deb,fac_fin, '');
 
      StTexte := Format(xlib +' provisoire n° %d , Voulez-vous la consulter ?',[fac_deb]);
    //  PGIInfoAf(StTexte,Titre);

      If (PGIAskAF(StTexte,titre)= mrYes) then
      begin

       numpiece := xnat+';'+DateToStr(TobEchApp.GetValue('AFA_DATEECHE')) +';'+xsouche+';'+IntTOStr(fac_deb)+';0;';
       AppelPiece([ numpiece, 'ACTION=MODIFICATION'],2);
       SetLength(Tmp,3) ;
       tmp [0]:=  'Piece';
       tmp [1]:=  'ACTION=CREATION';
       tmp [2]:=  'FAC';
       If (PGIAskAF('Voulez-vous Générer '+xlib2+' ?',titre)= mrYes) then
         Begin
         if AGlJaiLedroitFiche(tmp,3) then
          begin //mcd 03/07/2003
          st := xlib3+' sera généré(e) à la date du ';
          (* mcd 12/06/02 Zdate :=AGLLanceFiche('AFF','AFSAISDATE','','',
          'TITRE:Validation définitive de '+xlib3+';LIBELLE:'+st+';ZZDATE:'+gDateAppreciation+';');  *)
          Zdate:=AFLanceFiche_SaisieDate( 'TITRE:Validation définitive de '+xlib3+';LIBELLE:'+st+';ZZDATE:'+gDateAppreciation+';');
          if (zdate = '0') then
            begin
               PGIInfoAF('Traitement abandonné',titre);
               exit;
             end;


          TobComplete := TOB.Create ('Liste des pieces',NIL, -1);
          TobPiece := Tob.Create ('PIECE', TobComplete,-1);
  //        zecle:= '"FPR";'+'"GFP";"'+IntTOStr(fac_deb)+'";"0"';
          zecle:= '"'+xnat+'";'+'"'+xsouche+'";"'+IntTOStr(fac_deb)+'";"0"';

          TobPiece.selectDB (zecle, Nil, False);
          RegroupeLesPieces (TObCOmplete,xnat2,True, False,False,0,StrToDate(zdate));
          TobComplete.free;
          End;
         end;
      end;
    End;
  finally
  	TobRegAff.free;
    TobProfil.free;
    TOBArticles.Free;
  end;
END;

procedure TOF_APPRECIATION.TrtSpecifSolde ;
var Req,ztypgen: string;
    zz_pq,zz_p,zz_fr,zz_fo : double;
    DdateD,DdateF : TdateTime;
    zrep : integer;
    znat : string;
BEGIN
	// maj de toute l'activité en attente
    // PL le 05/03/02 : INDEX 4
  req := 'UPDATE ACTIVITE  SET act_etatvisafac="VIS",act_activiterepris="FAC",act_numapprec="'+
  inttostr(OffSetApp+TobEchApp.GetValue('AFA_NUMECHE'))+'"'+' WHERE act_affaire="' + gCodeAffaire +
  '" AND ACT_DATEACTIVITE <= "'+usdatetime(StrToDate(gDateAppreciation))+
  '" AND act_activiterepris="F" AND act_etatvisafac <> "VIS" AND act_etatvisa="VIS"'  ;

  ExecuteSQL(req);
  // afafichage du montant à facturer
	RecupMontantLignesvisa;
  //reprise des factures d'acompte

// PL le 05/03/02 : INDEX 4
  req := 'UPDATE PIECE  SET gp_facreprise ='
  +IntToStr(OffSetApp+TobEchApp.GetValue('AFA_NUMECHE'))
  + ' WHERE gp_affaire="' + gCodeAffaire + '" AND  gp_naturepieceg ="FAC" AND gp_facreprise < 0 AND GP_GENERAUTO = "FOR"';
  ExecuteSQL(req);

  ztypgen:='REPR';
  zrep :=  TobEchApp.GetValue('AFA_NUMECHE');
  DDateD := idate1900;
  DDateF := idate2099;
  zz_pq := 0;zz_p :=0; zz_fr := 0;zz_fo := 0;
  znat := '("FAC")';
  RecupTobLignePiece ( true, false, gCodeAffaire, znat,ztypgen,zrep,DDateD,DDateF,zz_pq,zz_p,zz_fr,zz_fo) ;

  if THEdit(GetControl('ZRACPTE')) <> nil then
   begin  // n'existe que sur fiche AfAppreComplete
   SetControlText('ZRACPTE',strf00(zz_p+zz_fo+zz_fr,DEV.decimale));
//PQ  SetControlText('ZRACPTE_PQ',strf00(zz_pq,DEV.decimale));
   SetControlText('ZRACPTE_P',strf00(zz_p,DEV.decimale));
   SetControlText('ZRACPTE_FR',strf00(zz_fr,DEV.decimale));
   SetControlText('ZRACPTE_FO',strf00(zz_fo,DEV.decimale));
   end;


  // Cumul En COurs Net
	CumECNET;
 	AfficheEcranSituation;

	ControleSaisie;


END;

//******************************************************************
//           SAISIE,  Ecran...
//******************************************************************

Procedure  TOF_APPRECIATION.ControleSaisie;
var zone,zone_p,zone_fr,zone_fo : double;
		cumfac,cumbon : double;
BEGIN

	cumbon:=0 ;

// controle sur ligne a facturer
  zone:= valeur(GetControlText('ZAFAC'));
  zone_p:=  valeur(GetControlText('ZAFAC_P'));
  zone_fr:=  valeur(GetControlText('ZAFAC_FR'));
  zone_fo:=  valeur(GetControlText('ZAFAC_FO'));

	if (zone = 0.0) and (  (zone_p <> 0.0)  or (zone_fr <> 0.0) or (zone_fo <> 0.0)) then
  begin
  	zone := arrondi((zone_p + zone_fr + zone_fo),DEV.decimale);
    SetControlText('ZAFAC',strf00(zone,DEV.decimale));
  end;
  if (zone <> 0.0) and (zone_p = 0.0)  and (zone_fr = 0.0) and (zone_fo = 0.0) then
  begin
  	zone_p := zone;
    SetControlText('ZAFAC_P',strf00(zone_p,DEV.decimale));
  end;

	cumfac := zone;

  if (arrondi(zone,DEV.decimale) <> arrondi((zone_p + zone_fr + zone_fo),DEV.decimale)) then
  Begin
  	AfficheError  := false;
  	LastError := 1;
  	LastErrorMsg :=  TexteMsgAffaire[1];
		PGIInfoAf(TexteMsgAffaire[1],titre);
    Erreur:=(LastError<>0);
    setfocuscontrol('ZAFAC');
    exit;
  End;

// controle sur ligne Boni/mali 1
	if (BM.bm1) then
  Begin
    zone:= valeur(GetControlText('ZABON1'));
    zone_p:=  valeur(GetControlText('ZABON1_P'));
    zone_fr:=  valeur(GetControlText('ZABON1_FR'));
    zone_fo:=  valeur(GetControlText('ZABON1_FO'));
    if (zone = 0.0) and (  (zone_p <> 0.0)  or (zone_fr <> 0.0) or (zone_fo <> 0.0)) then
    begin
      zone := arrondi((zone_p + zone_fr + zone_fo),DEV.decimale);
      SetControlText('ZABON1',strf00(zone,DEV.decimale));
    end;
    if (zone <> 0.0) and (zone_p = 0.0) and (zone_fr = 0.0) and (zone_fo = 0.0) then
    begin
      zone_p := zone;
      SetControlText('ZABON1_P',strf00(zone_p,DEV.decimale));
    end;

    if (arrondi(zone,DEV.decimale) <> arrondi((zone_p + zone_fr + zone_fo),DEV.decimale)) then
    Begin
      AfficheError  := false;
      LastError := 1;
      LastErrorMsg :=  TexteMsgAffaire[3];
      PGIInfoAf(TexteMsgAffaire[3],titre);
      Erreur:=(LastError<>0);
      setfocuscontrol('ZABON1');
      exit;
    End;

    cumbon := zone;
  End;

  // controle sur ligne Boni/mali 1
	if (BM.bm2) then
  Begin
    zone:= valeur(GetControlText('ZABON2'));
    zone_p:=  valeur(GetControlText('ZABON2_P'));
    zone_fr:=  valeur(GetControlText('ZABON2_FR'));
    zone_fo:=  valeur(GetControlText('ZABON2_FO'));
    if (zone = 0.0) and (  (zone_p <> 0.0)  or (zone_fr <> 0.0) or (zone_fo <> 0.0)) then
    begin
      zone := arrondi((zone_p + zone_fr + zone_fo),DEV.decimale);
      SetControlText('ZABON2',strf00(zone,DEV.decimale));
    end;
    if (zone <> 0.0) and (zone_p = 0.0)  and (zone_fr = 0.0) and (zone_fo = 0.0) then
    begin
      zone_p := zone;
      SetControlText('ZABON2_P',strf00(zone_p,DEV.decimale));
    end;
  if (arrondi(zone,DEV.decimale) <> arrondi((zone_p + zone_fr + zone_fo),DEV.decimale)) then
    Begin
      AfficheError  := false;
      LastError := 1;
      LastErrorMsg :=  TexteMsgAffaire[3];
      PGIInfoAf(TexteMsgAffaire[3],titre);
      Erreur:=(LastError<>0);
      setfocuscontrol('ZABON2');
      exit;
    End;
    cumbon := cumbon + zone;
  End;

  if (cumbon = 0) and (cumfac = 0) and (GetControlText('AFA_ETATVISA') = 'VIS')  then
    Begin
      AfficheError  := false;
      LastError := 1;
      LastErrorMsg :=  TexteMsgAffaire[4];
      PGIInfoAf(TexteMsgAffaire[4], titre);
      Erreur:=(LastError<>0);
      setfocuscontrol('ZAFAC');
      exit;
    End;

  if (GetControlText('AFA_ETATVISA') = '') then
    begin
      AfficheError := false;
      LastError := 1;
      LastErrorMsg :=  TexteMsgAffaire[8];
      PGIInfoAf (TexteMsgAffaire[8], titre);
      Erreur := (LastError <> 0);
      setfocuscontrol('AFA_ETATVISA');
      exit;
    end;

  //traitement specif solde si prochain ec différent de zéro
  if not(EcrSimple) and TCheckBox(GetControl('CHSOLD')).Checked    then
 		SoldeCreatBoniMali;

END;

Procedure  TOF_APPRECIATION.ModifMontAFacturer;
  Var
  zres, zpre, zfou, zfr, zcal : double;
BEGIN
  if not (EcrSimple) and not (TCheckBox (GetControl ('CHACPT')).Checked) then
    exit;

  if (GetControlText ('AFA_ETATVISA') = 'VIS') then
    exit;

  zres := valeur (GetControlText('ZAFAC'));
  //zpre  := valeur(GetControlText('ZECNET_P'));
  zfou := valeur (GetControlText('ZECNET_FO'));
  zfr := valeur (GetControlText('ZECNET_FR'));

  // je consomme dans l'ordre frais/four et prestation

  zcal := zres - zfr;
  if (zcal < 0) then
    begin
      zfr := zres;
      zpre := 0;
      zfou := 0;
    end
  else
    begin
      zres := zcal;
      zcal := zres - zfou;
      if (zcal < 0) then
        begin
          zfou := zres;
          zpre := 0;
        end
      else
        begin
          zpre := zcal;
        end
    end;

  SetControlText ('ZAFAC', strf00 (zfr + zpre + zfou, DEV.decimale));
  SetControlText ('ZAFAC_P', strf00 (zpre, DEV.decimale));
  SetControlText ('ZAFAC_FO', strf00 (zfou, DEV.decimale));
  SetControlText ('ZAFAC_FR', strf00 (zfr, DEV.decimale));
END;


procedure TOF_APPRECIATION.AcceslignesVisa(num : integer);
var zpar: string;
Begin
zpar := 'AFA_AFFAIRE:' + gCodeAffaire + ';ACT_DATEACTIVITE:'+gDateDebExCabinet+
';ACT_DATEACTIVITE_:'+gDateAppreciation+';AFA_REPRISEACTIV:TOU;'+
'NUMAPPREC:'+IntToStr(OffSetApp+TobEchApp.GetValue('AFA_NUMECHE'))+
';ACT_ETATVISA:VIS;ACT_TYPEACTIVITE:TOU;ACT_ETATVISAFAC:ATT;';

//If JaiLeDroitFiche('ACTIVITEMUL',ActionToString(TypeAction)) then
// mc d12/06/02AGLLanceFiche('AFF','AFACTIVITEGLO_MUL','','',zpar);
AFLanceFiche_Modif_Activite(Zpar);
// Alimentation du montant à facturer
RecupMontantLignesvisa;

Controlesaisie;// pour cohérence maj non synchro


END;

procedure TOF_APPRECIATION.AppelProd (typ : string);
  var zpar : string;
      dateDeb : TDateTime;
Begin
  if (typ = 'P')   then // acces à la prod
    begin
      if (GetParamSoc ('SO_AFAPPANOUVEAU')= true) then
        begin
          dateDeb := strtodate (gDateAppreciationPrec);
          if (dateDeb <> idate1900) then dateDeb := dateDeb + 1;
          zpar := 'AFA_AFFAIRE:' + gCodeAffaire + ';ACT_DATEACTIVITE:' + datetostr(dateDeb) + ';ACT_DATEACTIVITE_:' + gDateAppreciation
                  + ';AFA_REPRISEACTIV:TOU;ACT_ETATVISA="VIS";MONOSEL:;'
        end
      else
        zpar := 'AFA_AFFAIRE:' + gCodeAffaire + ';ACT_DATEACTIVITE:' + gDateDebExCabinet + ';ACT_DATEACTIVITE_:' + gDateAppreciation
            + ';AFA_REPRISEACTIV:TOU;ACT_ETATVISA="VIS";MONOSEL:;'
    end
  else
  if (typ = 'A')   then // acces à l'AN
    zpar := 'AFA_AFFAIRE:' + gCodeAffaire + ';ACT_DATEACTIVITE:' + DateToStr (idate1900)
            + ';ACT_DATEACTIVITE_:' + gDateDebExCabinet_1 + ';AFA_REPRISEACTIV:TOU;ACT_ETATVISA="VIS";MONOSEL:;'
  else
  if (typ = 'T')   then // acces à la totalité
    zpar := 'AFA_AFFAIRE:' + gCodeAffaire + ';ACT_DATEACTIVITE:' + DateToStr (idate1900)
            + ';ACT_DATEACTIVITE_:' + gDateAppreciationPrec + ';AFA_REPRISEACTIV:TOU;ACT_ETATVISA="VIS";MONOSEL:;';

  //If JaiLeDroitFiche('ACTIVITEMUL',ActionToString(TypeAction)) then
  AFLanceFiche_Mul_Activite (zpar);

END;

// appel liste de sfactures : typ = F facture A Acompte  T toutes
procedure TOF_APPRECIATION.AppelFact (typ: string; periode : string);
var zpar,zwhe,req: string;
      dateDeb : TDateTime;
Begin
zwhe := '';
Req := FabriqueReqPiece (typ, gConsultation);
// PL 16/12/02 : appreciation simple : on test qu'il y a qq chose dans req pour alimenter le zwhere...
if (Req <> '') then
  zwhe := 'XXWHERE:('+REQ+')';

// mcd 31/05/02 zpar := 'GP_AFFAIRE:'+gCodeAffaire+';GP_DATEPIECE:'+DateToStr(idate1900)+';GP_DATEPIECE_:'+gDateAppreciation+';'
if (periode = 'E') then
  begin
    if (GetParamSoc ('SO_AFAPPANOUVEAU')= true) then
        begin
          dateDeb := strtodate (gDateAppreciationPrec);
          if (dateDeb <> idate1900) then dateDeb := dateDeb + 1;
          zpar := 'GP_AFFAIRE:' + gCodeAffaire + ';GP_DATEPIECE:' + datetostr(dateDeb) +';GP_DATEPIECE_:'+gDateAppreciation+';'+zwhe
        end
    else
      zpar := 'GP_AFFAIRE:' + gCodeAffaire + ';GP_DATEPIECE:'+DateToStr(StoDateDeb)+';GP_DATEPIECE_:'+gDateAppreciation+';'+zwhe;
  end
else
if (periode = 'A') then
  zpar := 'GP_AFFAIRE:' + gCodeAffaire + ';GP_DATEPIECE:'+DateToStr(idate1900)+';GP_DATEPIECE_:'+gDateAppreciationPrec+';'+zwhe;

AFLanceFiche_PiecePourUneAffaire (zpar);

END;


// Alimentation du montant çà facturer
Procedure TOF_APPRECIATION.RecupMontantLignesvisa;
var req : string;
		zpre,zfra,zfou,zqte: Double;

BEGIN
// PL le 06/03/02 : INDEX 6
	req := 'SELECT SUM(ACT_TOTVENTE) AS ACT_TOTVENTE,SUM(ACT_QTE) AS ACT_QTE , ACT_TYPEARTICLE '+
  'FROM ACTIVITE ' +
  ' WHERE ACT_AFFAIRE="' + gCodeAffaire +
  '"  AND (ACT_NUMAPPREC = "' + inttostr (TobEchApp.GetValue ('AFA_NUMECHE')) + '"' +
  ' OR ACT_NUMAPPREC = "' + inttostr (OffSetApp + TobEchApp.GetValue ('AFA_NUMECHE')) + '") AND ACT_ETATVISAFAC = "VIS"' +
  ' GROUP BY ACT_TYPEARTICLE';

  zqte := 0; zpre := 0; zfra := 0; zfou := 0;
	RecupActiviteDansTOB (Req, zqte, zpre, zfra, zfou);

  SetControlText ('ZAFAC', strf00 (zpre + zfra + zfou, DEV.Decimale));
  SetControlText ('ZAFAC_P', strf00 (zpre, DEV.Decimale));
//PQ  SetControlText('ZAFAC_PQ',strf00(zqte,DEV.Decimale));
  SetControlText ('ZAFAC_FR', strf00 (zfra, DEV.Decimale));
  SetControlText ('ZAFAC_FO', strf00 (zfou, DEV.Decimale));

END;

Procedure TOF_APPRECIATION.RecupActivitedansTOB (Req : string; var zqte, zpre, zfra, zfou : Double);
  Var
    QQ : Tquery;
    TobAct, TobDet : TOB;
    wi : integer;
BEGIN
TobAct := TOB.Create ('lignes d''activité visées ', Nil, -1);
try
  QQ := nil;
  Try
    QQ := OpenSQL (Req, true);
    If Not QQ.EOF then TobAct.LoadDetailDB ('ACTIVITE', '', '', QQ, True);
  Finally
    Ferme (QQ);
  End;

  For wi := 0  to  TobAct.detail.count - 1 do
    Begin
      TobDet := TobAct.detail[wi];
      if TobDet.GetValue('ACT_TYPEARTICLE') = 'PRE' then
        begin
          zqte := TobDet.GetValue('ACT_QTE');
          zPre := TobDet.GetValue('ACT_TOTVENTE');
        end;
      if TobDet.GetValue('ACT_TYPEARTICLE') = 'MAR' then
        zfou := TobDet.GetValue('ACT_TOTVENTE');
      if TobDet.GetValue('ACT_TYPEARTICLE') = 'FRA' then
        zfra := zfra+ TobDet.GetValue('ACT_TOTVENTE');
      if TobDet.GetValue('ACT_TYPEARTICLE') = 'POU' then
        zfra := zfra+ TobDet.GetValue('ACT_TOTVENTE');      //mcd 08/03/02 pour art %
      if TobDet.GetValue('ACT_TYPEARTICLE') = 'CTR' then
        zpre := zpre+ TobDet.GetValue('ACT_TOTVENTE');      //mcd 08/08/02 pour art CTR
    End;

  finally
    TobAct.free;
  end;

END;

// ALim du montant à facturer à partir de l'encours
procedure TOF_APPRECIATION.RepriseEncours;
Begin

  TobEchApp.PutValue('AFA_AFACTOT',valeur(GetControlText('ZECNET')));
  TobEchApp.PutValue('AFA_AFACTURERQTE',valeur(GetControlText('ZECNET_PQ')));
  TobEchApp.PutValue('AFA_AFACTURER',valeur(GetControlText('ZECNET_P')));
  TobEchApp.PutValue('AFA_AFACTFO',valeur(GetControlText('ZECNET_FO')));
  TobEchApp.PutValue('AFA_AFACTFR',valeur(GetControlText('ZECNET_FR')));

  AfficheMontDecision;

  Actualisation;

End;
(*
// Affiche du montant de l'anouveau
procedure TOF_APPRECIATION.RepriseANouveau;
var
  ZanPrec, ZanProd_p, ZanProd_fr, ZanProd_fo : double;
Begin
  ZanPrec := TobEchAppPrec.GetValue ('AFA_AFACTOT');

  SetControlText ('ZANPREC', strf00 (ZanPrec, DEV.decimale));
  SetControlText ('ZANPROD_P', strf00 (TobEchAppPrec.GetValue ('AFA_AFACTURER'), DEV.decimale));
  SetControlText ('ZANPROD_FR', strf00 (TobEchAppPrec.GetValue ('AFA_AFACTFR'), DEV.decimale));
  SetControlText ('ZANPROD_FO', strf00 (TobEchAppPrec.GetValue ('AFA_AFACTFO'), DEV.decimale));

End;
*)
//Calcul du Prochain En cours
procedure TOF_APPRECIATION.Actualisation;
Begin
		ControleSaisie;
		CalcEC ('ZECNET', 'ZAFAC', 'ZABON1', 'ZABON2', 'ZEC');
End;

Procedure TOF_APPRECIATION.CalcEC (zecnet, zafac, zbon1, zbon2, zec : string);
var zcum : double;
		zsuff, zracpt : string;
    wi : integer;
BEGIN
	zracpt := 'ZRACPTE';
	for wi := 1 to maxsuff do
  Begin
    zsuff := suff[wi];
    zcum := valeur (GetControlText (zecnet + zsuff)) - valeur (GetControlText (zafac + zsuff));

    zcum := zcum + (valeur (GetControlText (zbon1 + zsuff)) * BM.sens1) +
                   (valeur (GetControlText (zbon2 + zsuff)) * BM.sens2);

    if (not (EcrSimple)) and (TCheckBox (GetControl ('CHSOLD')).Checked) then
    begin
        zcum := zcum + valeur( GetControlText (zracpt + zsuff));
    end;


    if zcum <> 0 then
			SetControlText (zec + zsuff, strF00 (zcum, DEV.Decimale))
    else
    	SetControlText (zec + zsuff, '0.00');
  End;

  SetControlText ('ZEC_PQ', ''); // provisoire en attendant de gérer les qtés

END;

// initialisation diverses
// ind : indice de départ
// zinit avec RAZ
Procedure TOF_APPRECIATION.ModifZoneDeci (zone : string; ind : integer; zenable, zinit : boolean; zcolor : Tcolor);
var zcum : double;
		zsuff : string;
    wi : integer;
BEGIN
	for wi := ind to maxsuff do
  Begin
    zsuff := suff[wi];
    if zinit then
    Begin
      zcum := 0;
      SetControlText(zone+zsuff,strF00(zcum,DEV.Decimale));
    End;
    SetControlEnabled( zone+zsuff,zenable);
    SetControlProperty(zone+zsuff,'Color',zcolor);
  End;

END;


{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 16/10/2001
Modifié le ... :   /  /
Description .. : Ecran appréciation , initialistaion des zones de la partie
Suite ........ : décision  , ind : indice de départ de la zone à mettre à jour
Mots clefs ... : GIGA;APPRECIATION
*****************************************************************}
Procedure TOF_APPRECIATION.InitZoneDeci(zone : string;ind : integer);
var zcum : double;
		zsuff : string;
    wi : integer;
BEGIN
	zcum := 0;
	for wi := ind to maxsuff do
  Begin
    zsuff := suff[wi];
    SetControlText(zone+zsuff,strF00(zcum,DEV.Decimale));
  End;
END;


//  affichae en fonction du type de facture
procedure TOF_APPRECIATION.TrtTypeFacture;
var zinit,chg : boolean;
Begin

	chg := ChangementTypeFact;
  if chg then zinit := true else zinit := false;  // initialisation slt si changement de type

  if TCheckBox(GetControl('CHACPT')).Checked     then
  Begin
  	//zinit := true;
  	ModifZoneDeci('ZAFAC',1,true,zinit,clwindow);
    ModifZoneDeci('ZABON1',1,false,zinit,clbtnface);
    ModifZoneDeci('ZABON2',1,false,zinit,clbtnface);

    SetControlVisible('TZAFAC',false);
    SetControlVisible('TZMOINS',false);
    if THEdit(GetControl('ZRACPTE')) <> nil then
      begin  // n'existe que sur fiche AfAppreComplete
      SetControlVisible('TZRACPTE',false);
      SetControlVisible('ZRACPTE',false); SetCOntrolText('ZRACPTE',strf00(0,V_PGI.OkDecV));
      SetControlVisible('ZRACPTE_PQ',false); SetCOntrolText('ZRACPTE_PQ',strf00(0,V_PGI.OkDecQ));
      SetControlVisible('ZRACPTE_P',false); SetCOntrolText('ZRACPTE_P',strf00(0,V_PGI.OkDecV));
      SetControlVisible('ZRACPTE_FO',false); SetCOntrolText('ZRACPTE_FO',strf00(0,V_PGI.OkDecV));
      SetControlVisible('ZRACPTE_FR',false); SetCOntrolText('ZRACPTE_FR',strf00(0,V_PGI.OkDecV));
      end;
    SetControlEnabled('AFA_JUSTIFBONI',false); SetControlProperty('AFA_JUSTIFBONI','Color',clbtnface);
    SetControlEnabled('AFA_JUSTIFBONI2',false);SetControlProperty('AFA_JUSTIFBONI2','Color',clbtnface);
    SetControlVisible('BLETTR',false);
    SetFocusControl('ZAFAC');
		TobEchApp.PutValue('AFA_GENERAUTO','FOR');
    TypPrec := 'FOR';
   End;

   if TCheckBox(GetControl('CHSOLD')).Checked    then
   Begin
      SetControlVisible('BLETTR',false);
      SetControlVisible('TZRACPTE',true);
      SetControlVisible('TZMOINS',true);
      SetControlVisible('TZAFAC',true);
      SetControlVisible('ZRACPTE',true);
      SetControlVisible('ZRACPTE_PQ',true);
      SetControlVisible('ZRACPTE_P',true);
      SetControlVisible('ZRACPTE_FO',true);
      SetControlVisible('ZRACPTE_FR',true);

      //zinit := false;
      ModifZoneDeci('ZAFAC',2,true,zinit,clbtnface);
      ModifZoneDeci('ZABON1',1,false,zinit,clbtnface);
      ModifZoneDeci('ZABON2',1,false,zinit,clbtnface);

      SetControlEnabled('AFA_JUSTIFBONI',true);SetControlProperty('AFA_JUSTIFBONI','Color',clwindow);
      SetControlEnabled('AFA_JUSTIFBONI2',true); SetControlProperty('AFA_JUSTIFBONI2','Color',clwindow);

      SetFocusControl('ZAFAC');


      if (TobEchApp.GetValue('AFA_GENERAUTO') <> 'SOL') then
      Begin
        If (PGIAskAF('Confirmez-vous votre demande de facture de solde',titre)= mrYes) then
        Begin
        	TrtSpecifSolde
        End;
      End;
      TobEchApp.PutValue('AFA_GENERAUTO','SOL');
      TypPrec := 'SOL';
   End;

   if TCheckBox(GetControl('CHFAC')).Checked   then
   Begin
      SetControlVisible('BLETTR',true);
      SetControlVisible('TZRACPTE',false);
      SetControlVisible('TZMOINS',false);
      SetControlVisible('TZAFAC',false);
      SetControlVisible('ZRACPTE',false); SetCOntrolText('ZRACPTE',strf00(0,V_PGI.OkDecV));
      SetControlVisible('ZRACPTE_PQ',false); SetCOntrolText('ZRACPTE_PQ',strf00(0,V_PGI.OkDecQ));
      SetControlVisible('ZRACPTE_P',false); SetCOntrolText('ZRACPTE_P',strf00(0,V_PGI.OkDecV));
      SetControlVisible('ZRACPTE_FO',false); SetCOntrolText('ZRACPTE_FO',strf00(0,V_PGI.OkDecV));
      SetControlVisible('ZRACPTE_FR',false); SetCOntrolText('ZRACPTE_FR',strf00(0,V_PGI.OkDecV));

      //zinit := false;
      //metter ZAFAC en grise
      ModifZoneDeci('ZAFAC',2,true,zinit,clwindow);
      ModifZoneDeci('ZABON1',1,true,zinit,clwindow);
      ModifZoneDeci('ZABON2',1,true,zinit,clwindow);

      SetControlEnabled('AFA_JUSTIFBONI',true);SetControlProperty('AFA_JUSTIFBONI','Color',clwindow);
      SetControlEnabled('AFA_JUSTIFBONI2',true); SetControlProperty('AFA_JUSTIFBONI2','Color',clwindow);

      SetFocusControl('ZAFAC');

      TobEchApp.PutValue('AFA_GENERAUTO','ACT');
      TypPrec := 'ACT';
   End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 16/10/2001
Modifié le ... :   /  /    
Description .. : Appréciation : Changement du type de facture;
Suite ........ : solde,acompte ou partielle
Mots clefs ... : 
*****************************************************************}
Function TOF_APPRECIATION.ChangementTypeFact:boolean;
var TypSais : string;
BEGIN
 result := false;
 if TCheckBox(GetControl('CHACPT')).Checked     then TypSais := 'FOR';
 if TCheckBox(GetControl('CHSOLD')).Checked     then TypSais := 'SOL';
 if TCheckBox(GetControl('CHFAC')).Checked     	then TypSais := 'ACT';
 if (TypPrec <> '') and (TypPrec <> Typsais) and (TypPrec <> 'FOR')  then
 begin
 		result := true ;  // changement
    If (PGIAskAF(TexteMsgAffaire[5],titre)= mrYes) then
    Begin
      RAZApprec
    end
    else
    Begin
        if (TypPrec = 'FOR') then SetControlChecked('CHACPT', true);
        if (TypPrec = 'SOL') then SetControlChecked('CHSOLD', true);
        if (TypPrec = 'ACT') then SetControlChecked('CHFAC', true);
    End;
 end;
 if (TypPrec = 'FOR') and (TypSais<>'FOR')  then  //cas du forfait je remets simplement les lignes à 0
   InitZoneDeci('ZAFAC',1);
END;



// Acces à l'échéance et à l'affaire , puis analyse des données
function TOF_APPRECIATION.InfosGene : integer;
var  LibDevise,zlib,DatEch,St,tlib,libcli,lib1,lib2: string;
     TG : Tgroupbox;

Begin
Result := 0;
//
// Recup Param Boni/Mali
//
    // Remplissage structure Boni/Mali suivant paramétrage dans la table PARBONI
    RecupParamBoniMali(BM);
    // Adaptation de l'ecran par rapport au paramétrage
    GestionZoneEcranBM(1);
    GestionZoneEcranBM(2);

//
// Affichage bandeau des infos générales
//
    // Gestion de l'affaire et du tiers
   RemplirTobAffaire (gCodeAffaire, TobAff);
   libcli := RecupLibelleTiers (TOBAff.GetValue ('AFF_TIERS'));
   st := CodeAffaireAffiche (gCodeAffaire, ' ');


   tlib := format('%s / %s   %s',[libcli,st,TOBAff.GetValue('AFF_LIBELLE')]);
   SetcontrolText('TLIBCLIAFF',tlib);

   SetControlText('AFF_AFFAIRE',TOBAff.GetValue('AFF_AFFAIRE'));
   SetControlText('AFF_TIERS',TOBAff.GetValue('AFF_TIERS'));

   lib1:='';lib2:='';
   LibelleRessource( TOBAff.GetValue('AFF_RESPONSABLE'),lib1,lib2);
   st := format('%s %s',[lib1,lib2]);

   SetControlText('TAFF_RESPONSABLE_',st);


//
// Gestion des dates et remplissage des tob echeances
//
  Ech_Ok := RechEcheance (DatEch); // Recherche des différentes echeances et remplissage TOBEchApp et TOBEchAppPrec

  // PL le 08/10/03 : Pouvoir reboucler si on trouve que l'appréciation précédente n'est pas visée
  // sauf en consultation
  // Si il n'existe pas encore d'échéance à cette date
  // et si il existe une echeance précédente
  // Si on n'est pas en consultation et que l'echeance precedente n'est pas visée, on n'a pas le droit
  // de générer à nouveau une appréciation.
  if Not (ech_ok) then
    if (TOBEchAppPrec.GetValue('AFA_NUMECHE') <> 0) then
      if (TOBEchAppPrec.GetValue('AFA_ETATVISA') <> 'VIS') and (not Empilage) and (not gConsultation) then
        begin
          If (PGIAskAF ('Vous ne pouvez pas apprécier cette affaire : l''appréciation précédente n''est pas visée.' + chr(10) + 'Voulez-vous modifier l''appréciation précédente ?', titre) = mrYes) then
            Result := 1
          else
            Result := 2;
          exit;
        end;

  if (ech_ok) then
 	  gDateAppreciation := DatEch // Si on trouve une echeance dans l'intervalle de dates, c'est cette date qui fait foi
  else
    // Si on ne trouve pas d'echeance, on en cree temporairement une que l'on supprimera en cas de sortie par
    // la croix (annulation des modifications)
    Begin
      // PL le 11/08/03 : on a besoin de savoir si c'est la premiere fois qu'on cree cette echeance
      gbPremiereEcheance := true;

        // creation d'une echeance à la date gDateAppreciation
      CreerFactAff (TOBAff, strtodate(gDateAppreciation), 'APP');
      RechEcheance (DatEch);

      // MCd pour recharger la TobEchApp correctement avec l'éch créée précédemment.
      // voir si autre solution (sinon, pas OK la 1ere fois que l'on veut enrgt un enrgt + création enrgt à blanc
      // dans factaff car dans tobechapp, affaire et numech pas renseigné ==> ecriture sur enrgt blanc du mtt saisi.
      // 2eme fois OK car enrgt factaff créé par ligne précédente.
    End;

  if (EcrSimple) then
  Begin
  	TobEchApp.PutValue('AFA_GENERAUTO','FOR');
    TobEchApp.PutValue('AFA_APPRECIEE','X');
  End;

  if (TobEchAppPrec.GetValue('AFA_AFFAIRE') = '') then
  // on ne peux consulter de precedente, il n'y en a pas
    begin
      SetControlProperty('BANDER', 'Enabled', false);
      SetControlProperty('BANAPR', 'Enabled', false);
      SetControlProperty('BANREA', 'Enabled', false);
      SetControlProperty('BANFAC', 'Enabled', false);
      SetControlProperty('BANBON1', 'Enabled', false);
    end
  else
    if (gDateApprecAnneePrec = idate1900) then
      SetControlProperty('BANAPR', 'Enabled', false);

  // Date de debut de l'affaire
  gDateDebAffaire := TOBAff.GetValue('AFF_DATEDEBUT');

  // Date de debut cabinet et j-1
  gDateDebExCabinet := DateToStr (GetParamSoc ('SO_AFDATEDEBCAB'));
  gDateDebExCabinet_1:= DateToStr (Plusdate (StrToDate (gDateDebExCabinet), -1, 'J'));

  // Affichage de la date d'appreciation
  SetControlText('ACT_DATEACTIVITE_', gDateAppreciation);

	zlib := Format('Appréciation au  %s', [gDateAppreciation]);
  if (GetControl('FE_DATEAPPREC') = nil) then
    begin
      TG := TGroupBox (GetControl ('FE_APPREC'));
      TG.caption := zlib;
    end
  else
    SetControlText('FE_DATEAPPREC', zlib);

  // trt devise
  DEV.Code := TOBAff.GetValue('AFF_DEVISE');
  if (DEV.Code = '') then DEV.code := V_PGI.DevisePivot;
  GetInfosDevise(DEV) ;
  LibDevise := DEV.Libelle;

end;

Procedure  TOF_APPRECIATION.GestionZoneEcranBM(Num : Integer);
Var zvis : boolean;
    pres,four,frais,znum : string;
BEGIN
		znum := IntToStr(num);

		if (num =1) then
    Begin
    	zvis := BM.bm1;
      pres := BM.pres1;
      four := BM.four1;
      frais := BM.frais1;
    End
    else
    Begin
    	zvis := BM.bm2;
      pres := BM.pres2;
      four := BM.four2;
      frais := BM.frais2;
 		End;


   if (not(zvis)) or (GetParamSoc('SO_AFAPPAVECBM')= false)  then

// Gestion  ligne de Boni/Mali  NON active
    begin
      If GetControl ('TBON' + znum )<>Nil then
        SetControlVisible ('TBON' + znum, false);
      SetControlVisible ('ZBON' + znum, false);
      SetControlVisible ('ZBON' + znum + '_P', false);
      SetControlVisible ('ZBON' + znum + '_PQ', false);
      SetControlVisible ('ZBON' + znum + '_PQ1', false);
      SetControlVisible ('ZBON' + znum + '_FR', false);
      SetControlVisible ('ZBON' + znum + '_FO', false);

      If GetControl ('TABON' + znum)<>Nil then SetControlVisible ('TABON' + znum, false);
//      If GetControl ('TABON' + znum + '_')<>Nil then SetControlVisible ('TABON' + znum + '_', false);
      SetControlVisible ('ZANBON' + znum, false);
      SetControlVisible ('BANBON' + znum, false);
      SetControlVisible ('ZABON' + znum, false);
      SetControlVisible ('ZABON' + znum + '_P', false);
      SetControlVisible ('ZABON' + znum + '_PQ', false);
      SetControlVisible ('ZABON' + znum + '_PQ1', false);
      SetControlVisible ('ZABON' + znum + '_FR', false);
      SetControlVisible ('ZABON' + znum + '_FO', false);

      // PL le 15/01/02 pour V575
      if (num = 1) then
        SetControlVisible ('AFA_JUSTIFBONI', false)
      else
        SetControlVisible ('AFA_JUSTIFBONI' + znum, false);

      SetControlVisible ('BBON', false);

    End
   else
    Begin
     // Gestion  ligne de Boni/Mali  active
     	if (num = 1) then
        Begin
          If GetControl ('TBON' + znum )<>Nil then
            SetControlText ('TBON' + znum, BM.lib1);
          SetControlText ('TABON' + znum, BM.lib1);
//          SetControlText ('TABON' + znum + '_', BM.lib1);
//          SetControlProperty ('BANBON' + znum, 'hint', 'Consultation ' + BM.lib1 + ' et ' + BM.lib2);
        end
      else
        Begin
          If GetControl ('TBON' + znum )<>Nil then
            SetControlText ('TBON' + znum, BM.lib2);
          SetControlText ('TABON' + znum, BM.lib2);
//          SetControlText ('TABON' + znum + '_', BM.lib2);
          //SetControlProperty ('BANBON' + znum, 'hint', 'Consultation ' + BM.lib2);
        end;

      if (pres = '') then
      Begin
      	  SetControlVisible('ZBON'+znum+'_P',false);
          SetControlVisible('ZBON'+znum+'_PQ',false);
          SetControlVisible('ZBON'+znum+'_PQ1',false);
      	  SetControlVisible('ZABON'+znum+'_P',false);
          SetControlVisible('ZABON'+znum+'_PQ',false);
          SetControlVisible('ZABON'+znum+'_PQ1',false);
      End;

      if (four = '') then
      Begin
					SetControlVisible('ZABON'+znum+'_FO',false);
          SetControlVisible('ZBON'+znum+'_FO',false);
      End;
      if (frais = '') then
      Begin
					SetControlVisible('ZABON'+znum+'_FR',false);
          SetControlVisible('ZBON'+znum+'_FR',false);
      End;
    end

END;

// Acces à l'échéance de l'affaire
function TOF_APPRECIATION.RechEcheance (var DatEch : String) : boolean;
var
  zdebm, zfinm : Tdatetime;
  Q : Tquery;
  zwhere, req : string;
  ech_ok, echprec_ok : boolean;
BEGIN
  // On aurait pu se débrouiller autrement en chargeant une tob avec toutes les appréciations et en allant
  // chercher la dernière dans l'intervalle, l'avant dernière et celle de l'année précédente...
  // ça aurait évité 3 requêtes... à voir...

  ech_ok := false;
  echprec_ok := false;

// Lecture Echeance dans le mois (ca implique qu'on a 1 seule appréciation par mois)
  zdebm :=  DebutDeMois (StrToDate (gDateAppreciation));
  zfinm :=  FinDeMois (StrToDate (gDateAppreciation));

// PL le 06/03/02 : INDEX 3
  zwhere := 'AFA_TYPECHE = "APP" and AFA_AFFAIRE="' + gCodeAffaire + '" and AFA_DATEECHE >= "' +
  usdatetime (zdebm) + '" and AFA_DATEECHE <= "' + usdatetime (zfinm) + '"';
  // SELECT * : nombre de champ restreint et le nombre d'enreg est très réduit
  req := 'SELECT * FROM FACTAFF WHERE ' + zwhere + ' ORDER BY AFA_DATEECHE DESC';

  Q := nil;
  try
    Q := OpenSQL (req, True, 1);
    if not (Q.eof) then
      Begin
        TOBEchApp.SelectDB ('', Q);
        ech_ok := true;
        DatEch := TobEchApp.GetValue ('AFA_DATEECHE');
      End
    else
      Begin
        TobEchApp.InitValeurs (false);
        TobEchApp.PutValue ('AFA_PROFILGENER', TobAff.GetValue ('AFF_PROFILGENER'));
      End;
  Finally
    Ferme(Q) ;
  End;

// Recherche Echéance précédente
//  If (ech_ok) then
//    Begin
      Q := nil;
      try
        // PL le 06/03/02 : INDEX 3
        // On prend maintenant les echeances strictement inférieures à la date d'appréciation courante
        // et donc plus besoin de chercher le deuxième dans la liste mais le premier pour avoir l'echeance
        // précédente
        zwhere :=  'where AFA_TYPECHE ="APP" AND AFA_AFFAIRE="' + gCodeAffaire + '" AND AFA_DATEECHE < "'+
        usdatetime (strtodate (gDateAppreciation)) + '" ORDER BY AFA_DATEECHE DESC';
//        Req := 'SELECT AFA_NUMECHE,AFA_TYPECHE FROM FACTAFF ' + zWhere ;
        Req := 'SELECT * FROM FACTAFF ' + zWhere ;

        // PL le 09/01/02 pour optimiser le nombre d'enregistrements lus en eagl
        Q := OpenSQL (Req, True, 1);
        // Q := OpenSQL (Req, True, 2);
        //  Q:=OpenSQL(Req,True);

        if Not Q.EOF then
          begin
            TOBEchAppPrec.SelectDB ('', Q);
            gCodeAffaire := TobEchAppPrec.GetValue ('AFA_AFFAIRE');
            gDateAppreciationPrec := TobEchAppPrec.GetValue ('AFA_DATEECHE');
            SetControlText ('DATEAPPRECPREC', gDateAppreciationPrec);

(*            //Q.First;  // classées par ordre décroissant
            //     Q.next;   // je cherche la 2eme
            if not Q.eof then
              begin
                echprec_ok := true;
                zind := q.Fields[0].AsInteger;
                ztyp := q.Fields[1].AsString;
                if (ztyp = 'APP') Then
                  begin
                    // PL le 06/03/02 : INDEX 3
                    zwhere :=  'where AFA_TYPECHE="APP" and AFA_AFFAIRE="'+gCodeAffaire+'" And AFA_NUMECHE ="'+
                    IntToStr(zind)+'"';
                    // SELECT * : nombre de champ restreint et le nombre d'enreg est très réduit
                    Req := 'SELECT * FROM FACTAFF ' + zWhere ;
                    QQ := OpenSQL (Req, True);
                    TOBEchAppPrec.SelectDB ('', QQ);
                    gCodeAffaire := TobEchAppPrec.GetValue ('AFA_AFFAIRE');
                    Ferme (QQ);
                  end;
              end;      *)
          end;
      Finally
        Ferme(Q);
      end;

//    end;

// Recherche Echéance année précédente : inactif pour l'instant
{  If (echprec_ok) then
    begin
      Q := nil;
      try
        zdebm :=  PlusDate (DebutDeMois (StrToDate (gDateAppreciation)), -1, 'A');
        zfinm :=  PlusDate (FinDeMois (StrToDate (gDateAppreciation)), -1, 'A');

        // PL le 06/03/02 : INDEX 3
        zwhere := 'AFA_TYPECHE = "APP" and AFA_AFFAIRE="' + gCodeAffaire + '" and AFA_DATEECHE >= "' +
        usdatetime(zdebm) + '" and AFA_DATEECHE <= "' +   usdatetime (zfinm) + '" ORDER BY AFA_DATEECHE DESC';
        Req := 'SELECT AFA_DATEECHE,AFA_NUMECHE,AFA_TYPECHE FROM FACTAFF WHERE ' + zWhere ;

        Q := OpenSQL (Req, True, 1);

        if not Q.eof then
          begin
            gDateApprecAnneePrec := q.Fields[0].AsDateTime;
          end;

      Finally
        Ferme (Q);
      end;

    end;}

  result := ech_ok;
END;


// Alimentation des données propres à l'échéance
// num = 0 1er passage  <> 0 re-affichage

procedure TOF_APPRECIATION.InfosSituationPeriode;
  var
    zrep : integer;
    zz, zz_pq, zz_p, zz_fr, zz_fo, zz2_pq, zz2_p, zz2_fr, zz2_fo : Double;
    znom, zart, ztypgen, znat : string;
    DdateD, DdateF : TdateTime;
BEGIN
//
// recherche des infos   A nouveaux
//
  znom := 'ZAN';

  // Par defaut on prend l'intervalle de date entre idate1900 et la date de début cabinet moins un jour
  DdateD := Idate1900;
  DdateF := StrToDate (gDateDebExCabinet_1);

  zz := 0; zz_pq := 0; zz_p :=0; zz_fr := 0; zz_fo := 0;
  zart := '';

  // On va chercher la production entre les date de debut et de fin (les totaux par type d'article)
  RecupActiviteTob ('REA', gCodeAffaire, DDateD, DDateF, zz_pq, zz_p, zz_fr, zz_fo);

  // On alimente la tob situation
  AlimChampTobSituation (znom, zz, zz_pq, zz_p, zz_fr, zz_fo);
  Cumulchamp (znom);


//
// recherche des infos   Production
//
  znom := 'ZPROD';


  DdateD := StrToDate (gDateDebExCabinet);
  DdateF := StrToDate (gDateAppreciation);
  zz_pq := 0; zz_p :=0; zz_fr := 0; zz_fo := 0;

  // On va chercher la production entre les date de debut et de fin (les totaux par type d'article)
  RecupActiviteTob ('REA', gCodeAffaire, DDateD, DDateF, zz_pq, zz_p, zz_fr, zz_fo);

  // On alimente la tob situation
  AlimChampTobSituation (znom, zz, zz_pq, zz_p, zz_fr, zz_fo);
  Cumulchamp (znom);


//
// recherche des infos Bonis/Malis
//
  if (GetParamSoc ('SO_AFAPPAVECBM') = false) then
    Begin
      InitTobSituation ('ZBON1', 'ZBON2_FO');
    End;

  if (GetParamSoc ('SO_AFAPPAVECBM') = true) then
    Begin
      DdateD := StrToDate (gDateDebAffaire);
      DdateF := StrToDate (gDateAppreciation);
      zz_pq := 0; zz_p :=0; zz_fr := 0; zz_fo := 0;
      zz2_pq := 0; zz2_p :=0; zz2_fr := 0; zz2_fo := 0;
      RecupBoniMaliTob (true, gConsultation, 'BON', gCodeAffaire, DDateD, DDateF, TobEchApp.GetValue ('AFA_NUMECHE'),
                        zz_pq, zz_p, zz_fr, zz_fo, zz2_pq, zz2_p, zz2_fr, zz2_fo);

    if (BM.bm1) then
      Begin
        znom := 'ZBON1';
        AlimChampTobSituation (znom, zz, (zz_pq * BM.sens1), (zz_p * BM.sens1), (zz_fr * BM.sens1), (zz_fo * BM.sens1));
        Cumulchamp (znom);
      End;

    if (BM.bm2) then
      Begin
        znom := 'ZBON2';
        AlimChampTobSituation (znom, zz, zz2_pq * BM.sens2, zz2_p * BM.sens2, zz2_fr * BM.sens2, zz2_fo * BM.sens2);
        Cumulchamp (znom);
      End;
    End;

  // Cumul Production Nette
  CumPRNET;


//
// recherche des infos  Facturation
//
  znom := 'ZFACT';
  DdateD := StrToDate (gDateDebAffaire);
  DdateF := StrToDate (gDateAppreciation);
  zz_pq := 0;zz_p :=0; zz_fr := 0; zz_fo := 0;

  ztypgen := 'FACT';
  znat := '("FAC","FRE","AVC")';
  zrep :=  TobEchApp.GetValue ('AFA_NUMECHE');
  StoDateDEb := DDateD;  // mcd 31/05/02 pour avoir la même date de début si bouton voir facture
  RecupTobLignePiece (false, gConsultation, gCodeAffaire, znat, ztypgen, zrep, DDateD, DDateF, zz_pq, zz_p, zz_fr, zz_fo);

  AlimChampTobSituation (znom, zz, zz_pq, zz_p, zz_fr, zz_fo);
  Cumulchamp (znom);


  if not (EcrSimple) then
    Begin
      // Cumul En COurs Brut
      CumECBRU;

      // recherche des infos  Facturation  (acomptes)
      znom := 'ZACPT';
      ztypgen := 'ACPT';

      zrep :=  TobEchApp.GetValue('AFA_NUMECHE');
      zz_pq := 0; zz_p :=0; zz_fr := 0; zz_fo := 0;
      znat := '("FAC","AVC")';
      RecupTobLignePiece (false, gConsultation, gCodeAffaire, znat, ztypgen, zrep, DDateD, DDateF, zz_pq, zz_p, zz_fr, zz_fo);

      AlimChampTobSituation (znom, zz, zz_pq, zz_p, zz_fr, zz_fo);
      Cumulchamp (znom);
    End;


  // Cumul En COurs Net
    CumECNET;

//
// Dans le cas de l'appreciation par Solde, on n'a pas fini... il nous faut la même chose mais pour
// l'appréciation precedente
//
  if (GetParamSoc ('SO_AFAPPANOUVEAU') = true) then
    begin
    //
    // recherche des infos   A nouveaux avant dernière appréciation
    //
      znom := 'ZANEC';

      // Par defaut on prend l'intervalle de date entre idate1900 et la date d'appréciation précédente
      DdateD := Idate1900;
      DdateF := StrToDate (gDateAppreciationPrec);

      zz := 0; zz_pq := 0; zz_p :=0; zz_fr := 0; zz_fo := 0;
      zart := '';

      // On va chercher la production entre les date de debut et de fin (les totaux par type d'article)
      RecupActiviteTob ('REA', gCodeAffaire, DDateD, DDateF, zz_pq, zz_p, zz_fr, zz_fo);

      // On alimente la tob situation
      AlimChampTobSituation (znom, zz, zz_pq, zz_p, zz_fr, zz_fo);
      Cumulchamp (znom);


    //
    // recherche des infos   Production  avant dernière appréciation
    //
      znom := 'ZANECPROD';


      DdateD := StrToDate (gDateAppreciationPrec);
      if (DdateD <> idate1900) then DdateD := DdateD + 1;

      DdateF := StrToDate (gDateAppreciation);

      zz_pq := 0; zz_p :=0; zz_fr := 0; zz_fo := 0;

      // On va chercher la production entre les date de debut et de fin (les totaux par type d'article)
      RecupActiviteTob ('REA', gCodeAffaire, DDateD, DDateF, zz_pq, zz_p, zz_fr, zz_fo);

      // On alimente la tob situation
      AlimChampTobSituation (znom, zz, zz_pq, zz_p, zz_fr, zz_fo);
      Cumulchamp (znom);


    //
    // recherche des infos Bonis/Malis avant dernière appréciation
    //
      if (GetParamSoc ('SO_AFAPPAVECBM') = false) then
        Begin
          InitTobSituation ('ZANECBON1', 'ZANECBON2_FO');
        End;

      if (GetParamSoc ('SO_AFAPPAVECBM') = true) then
        Begin
          DdateD := StrToDate (gDateAppreciationPrec);
          if (DdateD <> idate1900) then DdateD := DdateD + 1;

          DdateF := StrToDate (gDateAppreciation);

          zz_pq := 0; zz_p :=0; zz_fr := 0; zz_fo := 0;
          zz2_pq := 0; zz2_p :=0; zz2_fr := 0; zz2_fo := 0;
          RecupBoniMaliTob (true, gConsultation, 'BON', gCodeAffaire, DDateD, DDateF, TobEchApp.GetValue ('AFA_NUMECHE'),
                            zz_pq, zz_p, zz_fr, zz_fo, zz2_pq, zz2_p, zz2_fr, zz2_fo);

        if (BM.bm1) then
          Begin
            znom := 'ZANECBON1';
            TobSitu.PutValue('ZANECBON1', zz_pq * (zz_p + zz_fr + zz_fo));
          End;

        if (BM.bm2) then
          Begin
            znom := 'ZANECBON2';
            TobSitu.PutValue ('ZANECBON2', zz2_pq * (zz2_p + zz2_fr + zz2_fo));
          End;
        End;

      // Cumul Production Nette
      CumANECPRNET;


    //
    // recherche des infos  Facturation avant dernière appréciation
    //
      znom := 'ZANECFACT';
      DdateD := StrToDate (gDateAppreciationPrec);
      if (DdateD <> idate1900) then DdateD := DdateD + 1;

      DdateF := StrToDate (gDateAppreciation);

      zz_pq := 0;zz_p :=0; zz_fr := 0; zz_fo := 0;

      ztypgen := 'FACT';
      znat := '("FAC","FRE","AVC")';
      zrep :=  TobEchApp.GetValue ('AFA_NUMECHE');
      StoDateDEb := DDateD;  // mcd 31/05/02 pour avoir la même date de début si bouton voir facture
      RecupTobLignePiece (false, gConsultation, gCodeAffaire, znat, ztypgen, zrep, DDateD, DDateF, zz_pq, zz_p, zz_fr, zz_fo);

      AlimChampTobSituation (znom, zz, zz_pq, zz_p, zz_fr, zz_fo);
      Cumulchamp (znom);

      // Cumul En COurs Net
        CumANECECNET;


    //
    // recherche des infos   A nouveaux
    //
      znom := 'ZANAN';

      // Par defaut on prend l'intervalle de date entre idate1900 et la date d'appréciation précédente
      DdateD := Idate1900;
      DdateF := StrToDate (gDateAppreciationPrec);

      zz := 0; zz_pq := 0; zz_p :=0; zz_fr := 0; zz_fo := 0;
      zart := '';

      // On va chercher la production entre les date de debut et de fin (les totaux par type d'article)
      RecupActiviteTob ('REA', gCodeAffaire, DDateD, DDateF, zz_pq, zz_p, zz_fr, zz_fo);

      // On alimente la tob situation
      AlimChampTobSituation (znom, zz, zz_pq, zz_p, zz_fr, zz_fo);
      Cumulchamp (znom);

    //
    // A Nouveau recherche des infos   Production
    //
      znom := 'ZANPROD';

      DdateD := idate1900;
      DdateF := StrToDate (gDateAppreciationPrec);
      zz_pq := 0; zz_p :=0; zz_fr := 0; zz_fo := 0;

      // On va chercher la production entre les date de debut et de fin (les totaux par type d'article)
      RecupActiviteTob ('REA', gCodeAffaire, DDateD, DDateF, zz_pq, zz_p, zz_fr, zz_fo);

      // On alimente la tob situation
      AlimChampTobSituation (znom, zz, zz_pq, zz_p, zz_fr, zz_fo);
      Cumulchamp (znom);


    //
    // A Nouveau recherche des infos Bonis/Malis
    //
      if (GetParamSoc ('SO_AFAPPAVECBM') = false) then
        Begin
          InitTobSituation ('ZANBON1', 'ZANBON2_FO');
        End;

      if (GetParamSoc ('SO_AFAPPAVECBM') = true) then
        Begin
          DdateD := idate1900;
          DdateF := StrToDate (gDateAppreciationPrec);
          zz_pq := 0; zz_p :=0; zz_fr := 0; zz_fo := 0;
          zz2_pq := 0; zz2_p :=0; zz2_fr := 0; zz2_fo := 0;
          RecupBoniMaliTob (false, gConsultation, 'BON', gCodeAffaire, DDateD, DDateF, TobEchAppPrec.GetValue ('AFA_NUMECHE'),
                            zz_pq, zz_p, zz_fr, zz_fo, zz2_pq, zz2_p, zz2_fr, zz2_fo);

        if (BM.bm1) then
          Begin
            znom := 'ZANBON1';
            //TobSitu.PutValue('ZANBON1', zz_pq * (zz_p + zz_fr + zz_fo));
            AlimChampTobSituation (znom, zz, (zz_pq * BM.sens1), (zz_p * BM.sens1), (zz_fr * BM.sens1), (zz_fo * BM.sens1));
            Cumulchamp (znom);
          End;

        if (BM.bm2) then
          Begin
            znom := 'ZANBON2';
            //TobSitu.PutValue ('ZANBON2', zz2_pq * (zz2_p + zz2_fr + zz2_fo));
            AlimChampTobSituation (znom, zz, zz2_pq * BM.sens2, zz2_p * BM.sens2, zz2_fr * BM.sens2, zz2_fo * BM.sens2);
            Cumulchamp (znom);
          End;
        End;

      // Cumul Production Nette Anouveau
      CumANPRNET;


    //
    // A Nouveau recherche des infos  Facturation
    //
      znom := 'ZANFACT';
      DdateD := idate1900;
      DdateF := StrToDate (gDateAppreciationPrec);
      zz_pq := 0;zz_p :=0; zz_fr := 0; zz_fo := 0;

      ztypgen := 'FACT';
      znat := '("FAC","FRE","AVC")';
      zrep :=  TobEchAppPrec.GetValue ('AFA_NUMECHE');
      RecupTobLignePiece (true, gConsultation, gCodeAffaire, znat, ztypgen, zrep, DDateD, DDateF, zz_pq, zz_p, zz_fr, zz_fo);

      AlimChampTobSituation (znom, zz, zz_pq, zz_p, zz_fr, zz_fo);
      Cumulchamp (znom);


      // Cumul En COurs Net
      CumANECNET;
    end;



  AfficheEcranSituation

END;




procedure TOF_APPRECIATION.InfosDecision(num :integer);
var
    st,znom,zgen : string;

BEGIN
		AfficheMontDecision ;
    SetControlChecked('CHACPT', false);
    SetControlChecked('CHFAC', false);
    SetControlChecked('CHSOLD', false);

    zgen :=  TobEchApp.GetValue('AFA_GENERAUTO');
    TypPrec :=   zgen;
    if (zgen = 'FOR')  or (zgen = '') then
    		SetControlChecked('CHACPT', true);
    if zgen = 'ACT' then
    		SetControlChecked('CHFAC', true);
    if zgen = 'SOL' then
    		SetControlChecked('CHSOLD', true);

//    TrtTypeFacture;      inutile car le fait de positionner le bouton ci dessus
// lance la fonction

  if GetParamSoc ('SO_GCCOMMENTAIRE') then
    begin
      SetControlText('AFA_LIBELLEMULTI1',TobEchApp.GetValue('AFA_LIBELLE1'));
      SetControlText('AFA_LIBELLEMULTI2',TobEchApp.GetValue('AFA_LIBELLE2'));
      SetControlText('AFA_LIBELLEMULTI3',TobEchApp.GetValue('AFA_LIBELLE3'));
    end
  else
    begin
      SetControlText('AFA_LIBELLE1',TobEchApp.GetValue('AFA_LIBELLE1'));
      SetControlText('AFA_LIBELLE2',TobEchApp.GetValue('AFA_LIBELLE2'));
      SetControlText('AFA_LIBELLE3',TobEchApp.GetValue('AFA_LIBELLE3'));
    end;

    SetControlText('AFA_PROFILGENER',TobEchApp.GetValue('AFA_PROFILGENER'));
    SetControlText('AFA_BLOCNOTE',TobEchApp.GetValue('AFA_BLOCNOTE'));
    SetControlText('AFA_JUSTIFBONI',TobEchApp.GetValue('AFA_JUSTIFBONI'));
    SetControlText('AFA_JUSTIFBONI2',TobEchApp.GetValue('AFA_JUSTIFBONI2'));

    SetControlText('AFA_ETATVISA',TobEchApp.GetValue('AFA_ETATVISA'));
    SetControlText('AFA_VISEUR',TobEchApp.GetValue('AFA_VISEUR'));
    znom := RechDom('ttUtilisateur',TobEchApp.GetValue('AFA_VISEUR'),False);
    st := '';
    if TobEchApp.GetValue('AFA_ETATVISA') = 'VIS' then
    	st :=  Format('Le %s par %s', [TobEchApp.GetValue('AFA_DATEVISA'),znom]);
    SetControlText('LIBVISA',st);

END;


Procedure TOF_APPRECIATION.CreatTobSituation;
Begin

   TobSitu.AddChampSup('ZAN',False) ;   TobSitu.AddChampSup('ZAN_PQ',False) ;
   TobSitu.AddChampSup('ZAN_P',False) ; TobSitu.AddChampSup('ZAN_FR',False) ;
   TobSitu.AddChampSup('ZAN_FO',False) ;

   TobSitu.AddChampSup('ZPROD',False) ;    TobSitu.AddChampSup('ZPROD_PQ',False) ;
   TobSitu.AddChampSup('ZPROD_P',False) ;    TobSitu.AddChampSup('ZPROD_FR',False) ;
   TobSitu.AddChampSup('ZPROD_FO',False) ;

   TobSitu.AddChampSup('ZBON1',False) ;     TobSitu.AddChampSup('ZBON1_PQ',False) ;
   TobSitu.AddChampSup('ZBON1_P',False) ;    TobSitu.AddChampSup('ZBON1_FR',False) ;
   TobSitu.AddChampSup('ZBON1_FO',False) ;

   TobSitu.AddChampSup('ZBON2',False) ;       TobSitu.AddChampSup('ZBON2_PQ',False) ;
   TobSitu.AddChampSup('ZBON2_P',False) ;     TobSitu.AddChampSup('ZBON2_FR',False) ;
   TobSitu.AddChampSup('ZBON2_FO',False) ;

   TobSitu.AddChampSup('ZPRNET',False) ;       TobSitu.AddChampSup('ZPRNET_PQ',False) ;
   TobSitu.AddChampSup('ZPRNET_P',False) ;    TobSitu.AddChampSup('ZPRNET_FR',False) ;
   TobSitu.AddChampSup('ZPRNET_FO',False) ;

   TobSitu.AddChampSup('ZFACT',False) ;        TobSitu.AddChampSup('ZFACT_PQ',False) ;
   TobSitu.AddChampSup('ZFACT_P',False) ;     TobSitu.AddChampSup('ZFACT_FR',False) ;
   TobSitu.AddChampSup('ZFACT_FO',False) ;

 //****
  if not(EcrSimple) then
  Begin
     TobSitu.AddChampSup('ZECBRU',False) ;      TobSitu.AddChampSup('ZECBRU_PQ',False) ;
     TobSitu.AddChampSup('ZECBRU_P',False) ;     TobSitu.AddChampSup('ZECBRU_FR',False) ;
     TobSitu.AddChampSup('ZECBRU_FO',False) ;

     TobSitu.AddChampSup('ZACPT',False) ;        TobSitu.AddChampSup('ZACPT_PQ',False) ;
     TobSitu.AddChampSup('ZACPT_P',False) ;      TobSitu.AddChampSup('ZACPT_FR',False) ;
     TobSitu.AddChampSup('ZACPT_FO',False) ;
  End;

   TobSitu.AddChampSup('ZECNET',False) ;       TobSitu.AddChampSup('ZECNET_PQ',False) ;
   TobSitu.AddChampSup('ZECNET_P',False) ;      TobSitu.AddChampSup('ZECNET_FR',False) ;
   TobSitu.AddChampSup('ZECNET_FO',False) ;


   InitTobSituation('ZAN','ZECNET_FO');


  if (GetParamSoc ('SO_AFAPPANOUVEAU') = true) then
    begin
      TobSitu.AddChampSup('ZANEC',False) ;   TobSitu.AddChampSup('ZANEC_PQ',False) ;
      TobSitu.AddChampSup('ZANEC_P',False) ; TobSitu.AddChampSup('ZANEC_FR',False) ;
      TobSitu.AddChampSup('ZANEC_FO',False) ;

      TobSitu.AddChampSup('ZANECPROD',False) ;     TobSitu.AddChampSup('ZANECPROD_PQ',False) ;
      TobSitu.AddChampSup('ZANECPROD_P',False) ;   TobSitu.AddChampSup('ZANECPROD_FR',False) ;
      TobSitu.AddChampSup('ZANECPROD_FO',False) ;

      TobSitu.AddChampSup('ZANECBON1',False) ;     TobSitu.AddChampSup('ZANECBON1_PQ',False) ;
      TobSitu.AddChampSup('ZANECBON1_P',False) ;   TobSitu.AddChampSup('ZANECBON1_FR',False) ;
      TobSitu.AddChampSup('ZANECBON1_FO',False) ;

      TobSitu.AddChampSup('ZANECBON2',False) ;     TobSitu.AddChampSup('ZANECBON2_PQ',False) ;
      TobSitu.AddChampSup('ZANECBON2_P',False) ;   TobSitu.AddChampSup('ZANECBON2_FR',False) ;
      TobSitu.AddChampSup('ZANECBON2_FO',False) ;

      TobSitu.AddChampSup('ZANECPRNET',False) ;       TobSitu.AddChampSup('ZANECPRNET_PQ',False) ;
      TobSitu.AddChampSup('ZANECPRNET_P',False) ;    TobSitu.AddChampSup('ZANECPRNET_FR',False) ;
      TobSitu.AddChampSup('ZANECPRNET_FO',False) ;

      TobSitu.AddChampSup('ZANECFACT',False) ;     TobSitu.AddChampSup('ZANECFACT_PQ',False) ;
      TobSitu.AddChampSup('ZANECFACT_P',False) ;   TobSitu.AddChampSup('ZANECFACT_FR',False) ;
      TobSitu.AddChampSup('ZANECFACT_FO',False) ;

      TobSitu.AddChampSup('ZANECECNET',False) ;       TobSitu.AddChampSup('ZANECECNET_PQ',False) ;
      TobSitu.AddChampSup('ZANECECNET_P',False) ;      TobSitu.AddChampSup('ZANECECNET_FR',False) ;
      TobSitu.AddChampSup('ZANECECNET_FO',False) ;

      InitTobSituation('ZANECPROD','ZANECECNET_FO');



      TobSitu.AddChampSup('ZANAN',False) ;   TobSitu.AddChampSup('ZANAN_PQ',False) ;
      TobSitu.AddChampSup('ZANAN_P',False) ; TobSitu.AddChampSup('ZANAN_FR',False) ;
      TobSitu.AddChampSup('ZANAN_FO',False) ;

      TobSitu.AddChampSup('ZANPROD',False) ;     TobSitu.AddChampSup('ZANPROD_PQ',False) ;
      TobSitu.AddChampSup('ZANPROD_P',False) ;   TobSitu.AddChampSup('ZANPROD_FR',False) ;
      TobSitu.AddChampSup('ZANPROD_FO',False) ;

      TobSitu.AddChampSup('ZANBON1',False) ;     TobSitu.AddChampSup('ZANBON1_PQ',False) ;
      TobSitu.AddChampSup('ZANBON1_P',False) ;   TobSitu.AddChampSup('ZANBON1_FR',False) ;
      TobSitu.AddChampSup('ZANBON1_FO',False) ;

      TobSitu.AddChampSup('ZANBON2',False) ;     TobSitu.AddChampSup('ZANBON2_PQ',False) ;
      TobSitu.AddChampSup('ZANBON2_P',False) ;   TobSitu.AddChampSup('ZANBON2_FR',False) ;
      TobSitu.AddChampSup('ZANBON2_FO',False) ;

      TobSitu.AddChampSup('ZANPRNET',False) ;       TobSitu.AddChampSup('ZANPRNET_PQ',False) ;
      TobSitu.AddChampSup('ZANPRNET_P',False) ;    TobSitu.AddChampSup('ZANPRNET_FR',False) ;
      TobSitu.AddChampSup('ZANPRNET_FO',False) ;

      TobSitu.AddChampSup('ZANFACT',False) ;     TobSitu.AddChampSup('ZANFACT_PQ',False) ;
      TobSitu.AddChampSup('ZANFACT_P',False) ;   TobSitu.AddChampSup('ZANFACT_FR',False) ;
      TobSitu.AddChampSup('ZANFACT_FO',False) ;

      TobSitu.AddChampSup('ZANECNET',False) ;       TobSitu.AddChampSup('ZANECNET_PQ',False) ;
      TobSitu.AddChampSup('ZANECNET_P',False) ;      TobSitu.AddChampSup('ZANECNET_FR',False) ;
      TobSitu.AddChampSup('ZANECNET_FO',False) ;

      InitTobSituation('ZANPROD','ZANECNET_FO');
    end;

END;




// Cumul  des quantité prest , frais , four pour une zone données
Procedure TOF_APPRECIATION.Cumulchamp(nom : string);
var ind : integer;
    zcum : double;
BEGIN
     ind := TobSitu.GetNumChamp(nom);
     zcum := TobSitu.Getvaleur(ind+2) +   TobSitu.Getvaleur(ind+3) + TobSitu.Getvaleur(ind+4);
     TobSitu.PutValeur(ind,zcum);
END;

 // Alim des zones de la tob
Procedure TOF_APPRECIATION.AlimChampTobSituation(nom :string;zz,zz_pq,zz_p,zz_fr,zz_fo : double);
var ind : integer;
BEGIN
     ind := TobSitu.GetNumChamp(nom);
     TobSitu.Putvaleur(ind+1,zz_pq);
     TobSitu.Putvaleur(ind+2,zz_p);
     TobSitu.Putvaleur(ind+3,zz_fr);
     TobSitu.Putvaleur(ind+4,zz_fo);
END;

// Cumul  du PRNET
Procedure TOF_APPRECIATION.CumPRNET;
var wi : integer;
    zcum : double;
    nom : string;
BEGIN
	for wi := 1 to maxsuff do
    Begin
      nom := suff[wi];
      zcum := TobSitu.GetValue('ZAN'+nom) + TobSitu.GetValue('ZPROD'+nom)  +
              (TobSitu.GetValue('ZBON1'+nom)*BM.sens1)+(TobSitu.GetValue('ZBON2'+nom)*BM.sens2);
      TobSitu.putValue('ZPRNET'+nom,zcum);
    End;
END;

// Cumul  du ANPRNET
Procedure TOF_APPRECIATION.CumANPRNET;
var wi : integer;
    zcum : double;
    nom : string;
BEGIN
	for wi := 1 to maxsuff do
    Begin
      nom := suff[wi];
      zcum := TobSitu.GetValue('ZANPROD'+nom)  +
              (TobSitu.GetValue('ZANBON1'+nom)*BM.sens1)+(TobSitu.GetValue('ZANBON2'+nom)*BM.sens2);
      TobSitu.putValue('ZANPRNET'+nom,zcum);
    End;
END;

Procedure TOF_APPRECIATION.CumANECPRNET;
var wi : integer;
    zcum : double;
    nom : string;
BEGIN
	for wi := 1 to maxsuff do
    Begin
      nom := suff[wi];
      zcum := TobSitu.GetValue ('ZANECPROD' + nom)  +
              (TobSitu.GetValue ('ZANECBON1' + nom) * BM.sens1) + (TobSitu.GetValue ('ZANECBON2' + nom)*BM.sens2);
      TobSitu.putValue ('ZANECPRNET' + nom, zcum);
    End;
END;

// Cumul  du PRNET
Procedure TOF_APPRECIATION.CumECBRU;
var wi : integer;
    nom : string;
    zcum : double;
BEGIN
	for wi := 1 to MaxSuff do
    Begin
      nom := suff[wi];
      zcum := TobSitu.GetValue ('ZPRNET' + nom) - TobSitu.GetValue ('ZFACT' + nom);
      TobSitu.putValue ('ZECBRU' + nom, zcum);
    End;
END;

// Cumul  du  Encours Net
Procedure TOF_APPRECIATION.CumECNET;
var wi : integer;
    nom : string;
    zcum : double;
BEGIN
	for wi := 1 to MaxSuff do
    Begin
      nom := suff[wi];

      if not(EcrSimple) then
        zcum := TobSitu.GetValue('ZECBRU'+nom) - TobSitu.GetValue('ZACPT'+nom)
      else
        zcum := TobSitu.GetValue('ZPRNET' + nom) - TobSitu.GetValue('ZFACT' + nom);

      TobSitu.putValue('ZECNET' + nom, zcum);
    End;
END;

Procedure TOF_APPRECIATION.CumANECNET;
var wi : integer;
    nom : string;
    zcum : double;
BEGIN
	for wi := 1 to MaxSuff do
    Begin
      nom := suff[wi];
      zcum := TobSitu.GetValue('ZANPRNET'+nom) - TobSitu.GetValue('ZANFACT'+nom);

      TobSitu.putValue ('ZANECNET' + nom, zcum);
    End;
END;

Procedure TOF_APPRECIATION.CumANECECNET;
var wi : integer;
    nom : string;
    zcum : double;
BEGIN
	for wi := 1 to MaxSuff do
    Begin
      nom := suff[wi];
      zcum := TobSitu.GetValue('ZANECPRNET'+nom) - TobSitu.GetValue('ZANECFACT'+nom);

      TobSitu.putValue ('ZANECECNET' + nom, zcum);
    End;
END;


Procedure TOF_APPRECIATION.InitTobSituation(nomdeb,nomfin : string);
var deb,fin,wi : integer;
    zcum : double;
BEGIN
     zcum := 0.0;
     deb := TobSitu.GetNumChamp(nomdeb);
     fin := TobSitu.GetNumChamp(nomfin);
     for wi := deb to fin do
         TobSitu.Putvaleur(wi,zcum);

END;

Procedure TOF_APPRECIATION.AfficheEcranSituation;
  var
  deb, fin, wi : integer;
	wnom : string;
BEGIN
//
// Affichage Situation A nouveau
//
  if (GetParamSoc ('SO_AFAPPANOUVEAU') = true) then
    begin
      SetControlText ('ZANREA', StrF00( TobSitu.GetValue ('ZANAN'), DEV.Decimale));
      SetControlText ('ZANFAC', StrF00( TobSitu.GetValue ('ZANFACT'), DEV.Decimale));
      SetControlText ('ZANBON1', StrF00( TobSitu.GetValue ('ZANBON1') - TobSitu.GetValue ('ZANBON2'), DEV.Decimale));
//      SetControlText ('ZANBON2', StrF00( TobSitu.GetValue ('ZANBON2'), DEV.Decimale));

      SetControlText ('ZANPREC', StrF00( TobSitu.GetValue ('ZANECNET'), DEV.Decimale));
      SetControlText ('ZANPREC_P', StrF00( TobSitu.GetValue ('ZANECNET_P'), DEV.Decimale));
      SetControlText ('ZANPREC_FR', StrF00( TobSitu.GetValue ('ZANECNET_FR'), DEV.Decimale));
      SetControlText ('ZANPREC_FO', StrF00( TobSitu.GetValue ('ZANECNET_FO'), DEV.Decimale));

      SetControlText ('ZPROD', StrF00( TobSitu.GetValue ('ZANECPROD'), DEV.Decimale));
      SetControlText ('ZPROD_P', StrF00( TobSitu.GetValue ('ZANECPROD_P'), DEV.Decimale));
      SetControlText ('ZPROD_FR', StrF00( TobSitu.GetValue ('ZANECPROD_FR'), DEV.Decimale));
      SetControlText ('ZPROD_FO', StrF00( TobSitu.GetValue ('ZANECPROD_FO'), DEV.Decimale));
      SetControlText ('ZFACT', StrF00( TobSitu.GetValue ('ZANECFACT'), DEV.Decimale));
      SetControlText ('ZFACT_P', StrF00( TobSitu.GetValue ('ZANECFACT_P'), DEV.Decimale));
      SetControlText ('ZFACT_FR', StrF00( TobSitu.GetValue ('ZANECFACT_FR'), DEV.Decimale));
      SetControlText ('ZFACT_FO', StrF00( TobSitu.GetValue ('ZANECFACT_FO'), DEV.Decimale));
    end;

//
// Affichage Situation Periode
//
  if (GetParamSoc ('SO_AFAPPANOUVEAU') = true) then
    deb := TobSitu.GetNumChamp ('ZECNET')
  else
    deb := TobSitu.GetNumChamp ('ZAN');

  fin := TobSitu.GetNumChamp ('ZECNET_FO');

  for wi := deb to fin do
    Begin
      wnom := Tobsitu.GetNomChamp (wi);
      SetControlText (wnom, StrF00( TobSitu.GetValue (wnom), DEV.Decimale));
    end;
END;




Procedure TOF_APPRECIATION.AfficheMontDecision;
var
  zcum,ztot : double;
BEGIN

  SetControlText('ZAFAC',strF00(TobEchApp.GetValue('AFA_AFACTOT'),DEV.Decimale));
  SetControlText('ZAFAC_PQ',StrF00(TobEchApp.GetValue('AFA_AFACTURERQTE'),DEV.Decimale));
  SetControlText('ZAFAC_P',StrF00(TobEchApp.GetValue('AFA_AFACTURER'),DEV.Decimale));
  SetControlText('ZAFAC_FR',StrF00(TobEchApp.GetValue('AFA_AFACTFR'),DEV.Decimale));
  SetControlText('ZAFAC_FO',StrF00(TobEchApp.GetValue('AFA_AFACTFO'),DEV.Decimale));

  zcum := TobEchApp.GetValue('AFA_ACPTEPR') +
  				TobEchApp.GetValue('AFA_ACPTEFR') +
          TobEchApp.GetValue('AFA_ACPTEFO');

  if THEdit(GetControl('ZRACPTE')) <> nil then
    begin  // n'existe que sur fiche AfAppreComplete
    SetControlText('ZRACPTE',StrF00(zcum,DEV.Decimale));

    SetControlText('ZRACPTE_PQ',StrF00(TobEchApp.GetValue('AFA_ACPTEQTE'),DEV.Decimale));
    SetControlText('ZRACPTE_P',StrF00(TobEchApp.GetValue('AFA_ACPTEPR'),DEV.Decimale));
    SetControlText('ZRACPTE_FR',StrF00(TobEchApp.GetValue('AFA_ACPTEFR'),DEV.Decimale));
    SetControlText('ZRACPTE_FO',StrF00(TobEchApp.GetValue('AFA_ACPTEFO'),DEV.Decimale));
    end;

  if BM.bm1 then
  Begin
		ztot := TobEchApp.GetValue('AFA_BONIMALI')+
  					TobEchApp.GetValue('AFA_BM1FO') 		+
            TobEchApp.GetValue('AFA_BM1FR');

    SetControlText('ZABON1',Strf00(ztot,DEV.Decimale));     // ??GMGMGM
    SetControlText('ZABON1_PQ',StrF00(TobEchApp.GetValue('AFA_BONIMALIQTE'),DEV.Decimale));
    SetControlText('ZABON1_P',StrF00(TobEchApp.GetValue('AFA_BONIMALI'),DEV.Decimale));
    SetControlText('ZABON1_FO',StrF00(TobEchApp.GetValue('AFA_BM1FO'),DEV.Decimale));
    SetControlText('ZABON1_FR',StrF00(TobEchApp.GetValue('AFA_BM1FR'),DEV.Decimale));
  End;

  if BM.bm2 then
  Begin
	ztot := TobEchApp.GetValue('AFA_BM2PR') +
  				TobEchApp.GetValue('AFA_BM2FO') +
          TobEchApp.GetValue('AFA_BM2FR');

   SetControlText('ZABON2',Strf00(ztot,DEV.Decimale));
    SetControlText('ZABON2_PQ',StrF00(TobEchApp.GetValue('AFA_BM2PRQTE'),DEV.Decimale));
    SetControlText('ZABON2_P',StrF00(TobEchApp.GetValue('AFA_BM2PR'),DEV.Decimale));
    SetControlText('ZABON2_FR',StrF00(TobEchApp.GetValue('AFA_BM2FR'),DEV.Decimale));
    SetControlText('ZABON2_FO',StrF00(TobEchApp.GetValue('AFA_BM2FO'),DEV.Decimale));

 	End;

END;


Procedure TOF_APPRECIATION.RecupActiviteTob (TypAct, CodeAff : String; DateD, DateF : TdateTime;
                                              var zz_pq, zz_p, zz_fr, zz_fo : double);
  Var
  req, wheredate, wherevis, zgroup : string;
BEGIN
// PL le 06/03/02 : INDEX 1
  wheredate := ' AND ACT_DATEACTIVITE <= "'+ usdatetime (DateF)+'"';
  wheredate := wheredate + ' AND ACT_DATEACTIVITE >= "'+ usdatetime(DateD)+'"';
  wherevis := 'and ACT_ETATVISA = "VIS" ';
  zgroup := ' group by ACT_TYPEARTICLE';

  Req := 'SELECT SUM(ACT_TOTVENTE) AS ACT_TOTVENTE, SUM(ACT_QTE) AS ACT_QTE, ACT_TYPEARTICLE';
  Req := Req + ' FROM ACTIVITE WHERE ACT_TYPEACTIVITE = "' + TypAct + '" AND ACT_AFFAIRE = "' + CodeAff + '"';
  Req := Req + wheredate + wherevis + zgroup;

  RecupActiviteDansTob (Req, zz_pq, zz_p, zz_fr, zz_fo);

END;

Procedure TOF_APPRECIATION.RecupBoniMaliTob(SansApprecCourante, consult:boolean;TypAct,CodeAff :String ;DateD,DateF:TdateTime; napp : integer;
	var zz_pq,zz_p,zz_fr,zz_fo,zz2_pq,zz2_p,zz2_fr,zz2_fo: double) ;
Var req, zgroup, whereapp, wheredate : string;
		QQ : Tquery;
    TobAct, TobDet : TOB ;
    wi : integer;
BEGIN
	TobAct := TOB.Create ('lignes de boni mali à créer', Nil, -1);

// PL le 06/03/02 : INDEX 1
  wheredate := ' AND ACT_DATEACTIVITE <= "' + usdatetime (DateF)+'"';
  wheredate := wheredate + ' AND ACT_DATEACTIVITE >= "' + usdatetime (DateD) + '"';

  if (consult) then
    Begin
      if (SansApprecCourante) then
        begin
          if (Ecrsimple) then
            whereapp := ' AND (ACT_NUMAPPREC < "' + inttostr (napp) + '")'
          else
            whereapp := ' AND ((ACT_NUMAPPREC < "' + inttostr (napp) + '" and ACT_ACTIVITEREPRIS<>"FAC")' +
                        ' or (ACT_NUMAPPREC <= "' + inttostr (napp) + '" and ACT_ACTIVITEREPRIS="FAC"))'
        end
    End
  else
  	whereapp := '';

  zgroup := ' group by ACT_ARTICLE';

  Req := 'SELECT SUM(ACT_TOTVENTE) AS ACT_TOTVENTE,SUM(ACT_QTE) as ACT_QTE ,ACT_ARTICLE ' ;
  Req := Req + ' from ACTIVITE WHERE ACT_TYPEACTIVITE = "' + TypAct + '" AND ACT_AFFAIRE = "' + CodeAff + '"';
  Req := Req + wheredate + whereapp + zgroup;
  QQ := OpenSQL (Req, true);

  Try
    If Not QQ.EOF then
      TobAct.LoadDetailDB ('ACTIVITE', '', '', QQ, True);
  Finally
    Ferme (QQ);
  End;

  For wi := 0  to  TobAct.detail.count - 1 do
    Begin
      TobDet := TobAct.detail[wi];
      if TobDet.GetValue('ACT_ARTICLE') = BM.pres1 then
      begin
        zz_pq := TobDet.GetValue('ACT_QTE');
        zz_p := TobDet.GetValue('ACT_TOTVENTE');
      end;
      if TobDet.GetValue('ACT_ARTICLE') = BM.four1 then
        zz_fo := TobDet.GetValue('ACT_TOTVENTE');
      if TobDet.GetValue('ACT_ARTICLE') = BM.Frais1 then
        zz_fr := TobDet.GetValue('ACT_TOTVENTE');
      if TobDet.GetValue('ACT_ARTICLE') = BM.pres2 then
      begin
        zz2_pq := TobDet.GetValue('ACT_QTE');
        zz2_p := TobDet.GetValue('ACT_TOTVENTE');
      end;
      if TobDet.GetValue('ACT_ARTICLE') = BM.four2 then
        zz2_fo := TobDet.GetValue('ACT_TOTVENTE');
      if TobDet.GetValue('ACT_ARTICLE') = BM.Frais2 then
        zz2_fr := TobDet.GetValue('ACT_TOTVENTE');
    End;

  TobAct.free;
END;

procedure TOF_APPRECIATION.RecupTobLignePiece(AvecFactApprec, consult:boolean;CodeAff,znat,typgene : String;zrep:Integer ;DateD,DateF:TdateTime; var zz_pq,zz_p,zz_fr,zz_fo : double) ;
Var req ,wheredate ,wheregene, wherePiece ,req1 ,typ: string;
    QQ : tquery;
    TobLigne,TobDet : TOB;
    wi : integer;
BEGIN
	TobLigne := TOB.Create('lignes de piece',Nil,-1);

 	wheredate := ' AND GL_DATEPIECE <= "' + usdatetime (DateF) + '"';
  wheredate := wheredate + ' AND GL_DATEPIECE  >= "' + usdatetime (DateD) + '"';

  // on ne doit pas réintégrer la facture émise par l'appréciation
  if not AvecFactApprec then
    wherePiece := ' AND NOT (GP_DATEPIECE = "' + usdatetime (DateF) + '" AND GP_FACREPRISE=' + inttostr (zrep) + ') ';

  wheredate := wheredate + wherePiece;

  Req1 := '';

  WhereGene := '';
  typ := '';

  // pour mazar je récupere les facture au forfait non reprise et/ou les autres
  // pour les autres clients, je récupére toutes les factures

  if (typgene = 'FACT') then typ := 'F';
  if (typgene = 'ACPT') then typ := 'A';
  if (typgene = 'REPR') then typ := 'R';
  Req := FabriqueReqPiece(typ,consult);

  if (req <> '') then
  	wheregene := ' AND ' + req;
// PL le 22/01/02 ne prenait pas la souche dans la jointure
  Req1 := 'LEFT OUTER JOIN PIECE ON GL_NATUREPIECEG=GP_NATUREPIECEG AND GL_SOUCHE=GP_SOUCHE AND GL_NUMERO=GP_NUMERO ';
// Fin PL 22/01/02

// PL le 15/01/02 pour V575
  Req := 'SELECT SUM(GL_QTEFACT) as GL_QTEFACT,sum(GL_TOTALHT) as  GL_TOTALHT,'
  +'GL_TYPEARTICLE from LIGNE '+Req1+ ' WHERE GL_AFFAIRE = "'+
  CodeAff+'" and GL_NATUREPIECEG in '+znat+' and GL_TYPELIGNE="ART" '+
  wheredate+wheregene+ 'GROUP BY GL_TYPEARTICLE';
  QQ := OpenSQL(Req,true);
  Try
    If Not QQ.EOF then TobLigne.LoadDetailDB('LIGNE','','',QQ,True);
  Finally
    Ferme(QQ);
  End;


   For wi:=0  to  TobLigne.detail.count-1 do
  Begin
    TobDet := TobLigne.detail[wi];
    if TobDet.GetValue('GL_TYPEARTICLE') = 'PRE' then
      begin
      zz_pq := TobDet.GetValue('GL_QTEFACT');
      zz_p := zz_p+TobDet.GetValue('GL_TOTALHT'); // PL le 15/01/02 pour V575
      end;
    if TobDet.GetValue('GL_TYPEARTICLE') = 'MAR' then
      zz_fo := TobDet.GetValue('GL_TOTALHT'); // PL le 15/01/02 pour V575
    if TobDet.GetValue('GL_TYPEARTICLE') = 'FRA' then
      zz_fr := zz_fr+TobDet.GetValue('GL_TOTALHT'); // PL le 15/01/02 pour V575
    if TobDet.GetValue('GL_TYPEARTICLE') = 'POU' then
      zz_fr := zz_fr+TobDet.GetValue('GL_TOTALHT'); // mcd 08/03/02 pour art POU
    if TobDet.GetValue('GL_TYPEARTICLE') = 'CTR' then
      zz_p := zz_p+TobDet.GetValue('GL_TOTALHT'); // mcd 31/05/02 pour art CTR

  End;

  TobLigne.free;

END;

Function TOF_APPRECIATION.FabriqueReqPiece(typ : string;consult : boolean) : string;
var   zdate : TdateTime;
			Req : string;
      zrep : integer;
BEGIN

  // pour mazar je récupere les facture au forfait non reprise et/ou les autres
  // pour les autres clients, je récupére toutes les factures
  zrep :=  TobEchApp.GetValue('AFA_NUMECHE');
  result := '';
  // PL le 16/12/02 : pour le cas du bogue chez FGA : incohérence entre la consultation de son appréciation et les éditions
  // Si on est en type F et en appréciation simple : on sort
  if (typ = 'F') and EcrSimple then  exit;

  if (typ = 'F') then  // facture + acompte repris
    Req := '((GP_GENERAUTO<>"FOR" and GP_FACREPRISE <'+ inttostr(zrep)+ ')'+
           ' or (GP_GENERAUTO="FOR" and GP_FACREPRISE > 0 and GP_FACREPRISE <'+ inttostr(zrep)+ '))';


  zdate := TobEchApp.GetValue('AFA_DATEECHE');
  zdate := PlusDate(zdate,(-10),'J');
  // si acpte et consul , je rajoute un test sur la date car si mon acompte
  // créé en (-5001) est repris en (5004) , je ne sais plus quand il a été créé
  // je suppose que s'il a été créé par 5001 sa date facture est < dateecheance
  if (typ = 'A')  then   // acompte non repris ou en cours de reprise
  Begin
  		if not(consult) then
  			req := 'GP_GENERAUTO ="FOR" and (GP_FACREPRISE<0  or GP_FACREPRISE = '+inttostr(zrep)+')'
      else
  			req := ' GP_GENERAUTO ="FOR" and '+ '((GP_FACREPRISE >= ' + inttostr(zrep)
        + ' and GP_DATEPIECE < "'+usdatetime(zdate)+'")'
        +' or'
        +'( GP_FACREPRISE > '+inttostr(zrep*(-1))+ ' and GP_FACREPRISE<0))'
  End;

  if typ = 'R'  then   // acompte en cours de reprise, avec  num reprise
  Begin
  		req := 'GP_GENERAUTO ="FOR" and GP_FACREPRISE='+inttostr(OffSetApp+zrep)
  End;
  result := req;

END;

Procedure TOF_APPRECIATION.AppelAppreciationPrecedente;
begin
if (TobEchAppPrec.GetValue('AFA_DATEECHE') = idate1900) then
  PGIInfoAf (TexteMsgAffaire[9], titre)
else
  AFLanceFiche_AppreciationNiv0 ('ZORI:SIM1;ZAFF:' + gCodeAffaire + ';ZDATAPP:' + datetostr (TobEchAppPrec.GetValue('AFA_DATEECHE')) + ';');
end;

Procedure TOF_APPRECIATION.AppelAppreciationAnneePrecedente;
begin

if (gDateApprecAnneePrec = idate1900) then
  PGIInfoAf (TexteMsgAffaire[10], titre)
else
  AFLanceFiche_AppreciationNiv0 ('ZORI:SIM1;ZAFF:' + gCodeAffaire + ';ZDATAPP:' + datetostr (gDateApprecAnneePrec) + ';');

end;

Procedure TOF_APPRECIATION.AppelModifBoniMali;
var
  Arg, RepEF : string;
begin

  Arg := 'AFFAIRE=' + gCodeAffaire + ';';
  Arg := Arg + 'NUM=' + IntToStr(TobEchApp.GetValue('AFA_NUMECHE')) + ';';
  Arg := Arg + 'DATE=' + datetostr (TobEchApp.GetValue('AFA_DATEECHE')) + ';';
  Arg := Arg + 'TIERS=' + TobEchApp.GetValue('AFA_TIERS') + ';';
  RepEF := AFLanceFiche_ModifBoni (Arg) ;

end;

Procedure TOF_APPRECIATION.CalculeAFacturer;
var
  fac_p, fac_fo, fac_fr, fac : double;
begin
  fac_p := 0;
  fac_fo := 0;
  fac_fr := 0;

  if (GetControlText ('ZAFAC_P') <> '') then
    fac_p := strtofloat (GetControlText ('ZAFAC_P'));
  if (GetControlText ('ZAFAC_FO') <> '') then
    fac_fo := strtofloat (GetControlText ('ZAFAC_FO'));
  if (GetControlText ('ZAFAC_FR') <> '') then
    fac_fr := strtofloat (GetControlText ('ZAFAC_FR'));

  fac := fac_p + fac_fo + fac_fr;

  SetControlText ('ZAFAC', strf00 (fac, DEV.decimale));

end;


procedure AGLAcceslignesVisa(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;

if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;

if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION(MaTOF).AcceslignesVisa(1) else exit;
end;

procedure AGLRepriseEnCours(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION(MaTOF).RepriseEnCours else exit;
end;

procedure AGLValideVisa(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION(MaTOF).ValideVisa
 else exit;
end;

procedure AGLActualisation(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION(MaTOF).Actualisation else exit;
end;

procedure AGLCalculEncours(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION(MaTOF).CalcEC('ZECNET','ZAFAC','ZABON1','ZABON2','ZEC') else exit;
end;

procedure AGLTrtTypeFacture(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION(MaTOF).TrtTypeFacture else exit;
end;

procedure AGLAppelProd(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION(MaTOF).AppelProd(Parms[1]) else exit;
end;

procedure AGLAppelFact(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION(MaTOF).AppelFact(Parms[1], Parms[2]) else exit;
end;

procedure AGLModifMontAFacturer(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION(MaTOF).ModifMontAFacturer else exit;
end;

Function AFLanceFiche_AppreciationNiv0 (Argument : string) : variant;
begin
  if (GetParamSoc ('SO_AFAPPANOUVEAU') = true) then
    result := AGLLanceFiche ('AFF', 'AFAPPRECSOLDE', '', '', argument)
  else
    result := AGLLanceFiche ('AFF', 'AFAPPREC', '', '', argument);
end;

Function AFLanceFiche_AppreciationCompl(Argument:string):variant;
begin
result:=AGLLanceFiche('AFF','AFAPRECOMPLETE','','',argument);
end;


procedure AGLAppelAppreciationPrecedente(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION(MaTOF).AppelAppreciationPrecedente else exit;
end;

procedure AGLAppelAppreciationAnneePrecedente(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION(MaTOF).AppelAppreciationAnneePrecedente else exit;
end;


procedure AGLAppelModifBoniMali (parms : array of variant; nb : integer);
var  F : TForm ;
     MaTOF  : TOF;
begin
F := TForm (Longint (Parms[0]));
if (F is TFVierge) then MaTOF := TFVierge (F).LaTOF else exit;
if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION (MaTOF).AppelModifBoniMali else exit;
end;

procedure AGLCalculeAFacturer (parms : array of variant; nb : integer);
var  F : TForm ;
     MaTOF  : TOF;
begin
F := TForm (Longint (Parms[0]));
if (F is TFVierge) then MaTOF := TFVierge (F).LaTOF else exit;
if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION (MaTOF).CalculeAFacturer else exit;
end;

procedure AGLAfficheSituation (parms : array of variant; nb : integer);
var  F : TForm ;
     MaTOF  : TOF;
begin
F := TForm (Longint (Parms[0]));
if (F is TFVierge) then MaTOF := TFVierge (F).LaTOF else exit;
if (MaTOF is TOF_APPRECIATION) then TOF_APPRECIATION (MaTOF).InfosSituationPeriode else exit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 30/07/2003
Modifié le ... :   /  /    
Description .. : Fct générique pour annuler appréciation non encore visée.
Suite ........ : Fait à partir de la fct d'origine SupprApprec qui était dans la
Suite ........ : classe.
Suite ........ : Annule activite
Suite ........ : piece
Suite ........ : factaff
Mots clefs ... : APPRECIATION;
*****************************************************************}
Procedure AnnulationApprec (affaire : string; NumEche : integer; Simulation : boolean);
Var
  req : string;
begin    // voir fct AnnulAPprec de UtofAfPieceProAnnu .. fct similaire
  try
    req := 'delete  from FACTAFF '
    + ' where afa_affaire="'+Affaire+'"'
    + '  and AFA_NUMECHE ="'+IntToStr(NumEche)+ '"';
    ExecuteSQL(req);

    if not (Simulation) then
      Begin   //mcd 30/07/03 ajout test sur type activite Réalisé.. sinon, remet à balnc la zone numapprec même pour le boni, qui nepeuvent plus être ensuite supprimé
        req := 'UPDATE ACTIVITE  SET act_etatvisafac="ATT",act_numapprec=0,act_activiterepris="F" '
        +' WHERE act_typeActivite="REA" and act_affaire="'+Affaire+'"'+ ' AND act_numapprec="'+ inttostr(NumEche)+'"';
        ExecuteSQL(req);

        // recherche ancien num apprec
        MajPieceAvecAncienNumApprec (Affaire,NumEche);

        //annulation boni
        req := 'delete from ACTIVITE  '
        +' where ACT_AFFAIRE="'+Affaire+'"'+ ' and ACT_NUMAPPREC ="'
        + inttostr(NumEche)+'"'
        + 'and ACT_TYPEACTIVITE = "BON"';
        ExecuteSQL(req);
      End;
  except
    V_PGI.IoError := oeUnknown;
  end;
end;



Initialization
  registerclasses ( [ TOF_APPRECIATION ] ) ;
  RegisterAglProc('AcceslignesVisa',TRUE,1,AGLAcceslignesVisa);
  RegisterAglProc('RepriseEnCours',TRUE,1,AGLRepriseEnCours);
  RegisterAglProc('Actualisation',TRUE,1,AGLActualisation);
  RegisterAglProc('CalculEncours',TRUE,1,AGLCalculEncours);
  RegisterAglProc('TrtTypeFacture',TRUE,1,AGLTrtTypeFacture);
  RegisterAglProc('AppelProd',TRUE,1,AGLAppelProd);
  RegisterAglProc('AppelFact',TRUE,2,AGLAppelFact);
  RegisterAglProc('ModifMontAFacturer',TRUE,0,AGLModifMontAFacturer);
  RegisterAglProc('AppelAppreciationPrecedente',TRUE,0,AGLAppelAppreciationPrecedente);
  RegisterAglProc('AppelAppreciationAnneePrecedente',TRUE,0,AGLAppelAppreciationAnneePrecedente);
  RegisterAglProc('AppelModifBoniMali',TRUE,0,AGLAppelModifBoniMali);
  RegisterAglProc('CalculeAFacturer',TRUE,0,AGLCalculeAFacturer);
  RegisterAglProc('AfficheSituation',TRUE,0,AGLAfficheSituation);
end.
