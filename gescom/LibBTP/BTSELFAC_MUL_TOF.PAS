{***********UNITE*************************************************
Auteur  ...... : LS
Créé le ...... : 28/03/2008
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSELFAC_MUL ()
            .. : Nouvelle gestion de la selection pour facturation
Mots clefs ... : TOF;BTSELFAC_MUL
*****************************************************************}
Unit BTSELFAC_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Grids,
     AglInit,
{$IFNDEF EAGLCLIENT}
     db,DBGrids,FE_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,MailEagl,
{$ENDIF}
		 Graphics,
     types,
		 Menus,
     HQry,
     HDB,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     UTOF,
     CalcOlegenericBTP,
     EntGC,
     HTB97,
     SaisUtil,
     AffaireUtil,
     splash,
     BTPREPSELFAC_MUL_TOF,
     UTofAfBaseCodeAffaire,uEntCommun,UtilTOBPiece ;

Type
  TmtFact = (TmtFacturation,TmtCloture);
  TResultS = (BTrSOk,BTrsNone,BtrsAnnul);
  TOF_BTSELFAC_MUL = Class (TOF_AFBASECODEAFFAIRE)
      procedure OnNew                    ; override ;
      procedure OnDelete                 ; override ;
      procedure OnUpdate                 ; override ;
      procedure OnLoad                   ; override ;
      procedure OnArgument (S : String ) ; override ;
      procedure OnDisplay                ; override ;
      procedure OnClose                  ; override ;
      procedure OnCancel                 ; override ;
      procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
    private
    	DateCloture : TDateTime;
    	fTypeFacturation : string;
      ModeTraitement : TmtFact;
    	MnzPiece : TmenuItem;
      TOBFacturable,TOBmemoFact,TOBAFACTURER,TOBReference,TheOldFac : TOB;
      Blance : TToolbarButton97;
      BEfface: TToolbarButton97;
      Fliste : THDbGrid;
      GS : THgrid;
      SqlBase,ColName : string;
      RefPieceSel : String;
      Avancement : boolean;
      ReajusteDossierfac : boolean;
      RecupSaisie : boolean;

      Domaine : THValcomboBox;
      //
      procedure SetEventsGridListe(Etat : boolean);
      procedure GetControls;
      procedure FlisteDblClick (Sender : TObject);
      function Chargefacturation(TOBFacture : TOB) : boolean;
      function WhereRefPieceSel(TOBL : TOB) : string;
      procedure ChargeLesDevisNonFacture(Affaire,TypeFac: string; NumSelected,Indice : integer);
      procedure NewFacturation (Naturepieceg,Souche : string; Numero,Indice : integer);
      procedure InitGrilleGS;
      procedure AfficheLagridGS;
      procedure TraiteChampAffaireGrid(Gr: THGrid; Nom: string; Col: integer);
      procedure SetEventsGridGS (etat : boolean);
      procedure FlipThisRow (Sender: TObject; ARow: Integer;var Cancel: Boolean);
      function  Existefac (TOBL : TOB): boolean;
      procedure SetEvents;
      procedure BlanceClick (Sender : Tobject);
      procedure BeffaceClick (Sender : Tobject);
      procedure EnregistreSelection;
      procedure Majmemorisation;
      procedure SupprimeMemo;
      procedure BchercheClick;
      procedure beforeSelection;
      function FindFirstSelected: TOB;
      function IsPieceSelectionnable(Arow: integer): boolean;
      function ExisteFacturationOld (var Affaire,Naturepieceg,Souche : string; var Numero,Indice : integer) : TResultS;
      procedure ChargeOldfacturation (Naturepieceg,Souche : string; var Numero,IndiceG : integer);
      function ExistDevis(cledoc: r_cledoc): boolean; overload;
      function ExistDevis(TOBL : TOB): boolean; overload;
    	function IsPresentDansFactureOld(var TheRefDevis : string;cledoc: r_cledoc): boolean;
      procedure VoirPiece (Sender : TObject);
      function MontantDevZero : boolean;
      procedure GSPostDrawCell (ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    	procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);
    	procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);
      procedure VerifDevisFromFacturation (TOBTMP : TOB);
    	function ControleEtatDevisPrinc(NaturepieceG, Souche: string; Numero,IndiceG: integer): boolean;
      function ConfirmationCloture : boolean;
      function ControleArticleEcartFact : boolean;
    	function AddSelection(TOBP: TOB; NumOrdre: integer;var MessageErreur: string): boolean;
      procedure MiseANiveauDesDocuments;
    //
      procedure ControleChamp(Champ, Valeur: String);
    procedure Definifacturation(NaturePieceG, SoucheG: string; Numero,IndiceG: integer);
    function AcompteNonRegle (Naturepieceg,Souche : string; numero,indiceg : integer) : boolean;
    //
  end ;


Implementation
uses factcomm,FactTOB,FactTvaMilliem,FactCOmmBTP,facture,UtilPgi,
     UtilArticle,ParamSoc,UTOFBTClotureDev,RECALCPIECE_RAP_TOF,
     DialogEx;

procedure TOF_BTSELFAC_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSELFAC_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSELFAC_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTSELFAC_MUL.OnLoad ;
begin
  Inherited ;
  GS.visible := false;
end ;

procedure TOF_BTSELFAC_MUL.OnArgument (S : String ) ;
Var x				: integer;
    Critere : String;
    ValMul	: String;
    ChampMul: String;
begin
  Inherited ;
  //
  ReajusteDossierfac := false;
  RecupSaisie := false;
  StArgument := S;
  //
  Repeat
  Critere:=uppercase(ReadTokenSt(stArgument)) ;
  if Critere<>'' then
      begin
      x:=pos('=',Critere);
      if x<>0 then
         begin
         ChampMul:=copy(Critere,1,x-1);
         ValMul:=copy(Critere,x+1,length(Critere));
         end
      else
         ChampMul := Critere;
      ControleChamp(ChampMul, ValMul);
      end;
  until Critere='';

  if ModeTraitement = TmtCloture then
  begin
  	DateCloture := V_PGI.DateEntree;
  	ecran.Caption := Traduirememoire('Cloture de facturation');
    UpdateCaption(ecran);
  end;
  if ReajusteDossierfac then
  begin
  	ecran.Caption := Traduirememoire('Réajustement dossier de facturation');
    UpdateCaption(ecran);
    TToolbarButton97 (GetControl('BACOMPTES')).Visible := false;
  end;

  TheOldFac := TOB.Create ('LA DERNIERE FACTURE EMISE',nil,-1);
  SqlBase := 'SELECT GP_AFFAIRE,GP_NATUREPIECEG,GP_DATEPIECE,GP_NUMPIECE,GP_PRENUMERO,GP_NUMERO,GP_SOUCHE,GP_INDICEG,GP_TIERS,'+
             'GP_ETABLISSEMENT,GP_REFINTERNE,GP_REPRESENTANT,GP_APPORTEUR,GP_DEVISE,GP_TOTALTTC,GP_TOTALHT,GP_ETATVISA,GP_VIVANTE,'+
             'GP_EDITEE,GP_AFFAIRE1,GP_AFFAIRE2,GP_AFFAIRE3 ,GP_AVENANT,GP_AFFAIREDEVIS,AFF_AFFAIRE1,AFF_AFFAIRE2,'+
             'AFF_AFFAIRE3 ,AFF_AFFAIRE0,AFF_AVENANT,AFF_AFFAIRE,AFF_ETATAFFAIRE, AFF_GENERAUTO,AFF_ADMINISTRATIF, AFF_MODELE,AFF_OKSIZERO,PRG_TYPERG '+
             'FROM PIECE '+
             'LEFT OUTER JOIN AFFAIRE on AFF_AFFAIRE=GP_AFFAIREDEVIS '+
             'LEFT OUTER JOIN PIECERG on GP_NATUREPIECEG=PRG_NATUREPIECEG AND GP_SOUCHE=PRG_SOUCHE AND GP_NUMERO=PRG_NUMERO AND GP_INDICEG=PRG_INDICEG ';
  //
  TOBFacturable := TOB.create ('LA LISTE DES DEVIS',nil,-1);
  TOBmemoFact   := TOB.Create ('LES MEMORISATIONS',nil,-1);
  TOBAFACTURER  := TOB.Create ('LES PIECES',nil,-1);
  //
  GetControls;
  SetEventsGridListe(True);  // sur Fliste
  SetEvents;
  InitGrilleGS;
  //
  GS.visible := false;

end ;

procedure TOF_BTSELFAC_MUL.ControleChamp(Champ : String;Valeur : String);
var St_Plus : string ;
begin
  if Champ = 'REAJUSTEAVANC' then
  begin
    ReajusteDossierfac := true;
  end;
  if Champ = 'CLOTURE' then
    ModeTraitement := TmtCloture
  else if Champ = 'MODIFICATION' then
    ModeTraitement := TmtFacturation;

  if Champ ='SOUSTRAITANCE' then SetControlText('XX_WHERE', ' AND SSTRAITE > 0');

  //
  if Champ='STATUT' then
  Begin
    if Valeur = 'APP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'W')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'W');
      SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("","W")');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Appel');
      SetControlText('TAFF_AFFAIRE', 'Appel');
    end
    else if Valeur = 'INT' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'I')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'I');
      SetControlText('XX_WHERE', ' AND ((SUBSTRING(GP_AFFAIRE,1,1)="I") OR (GP_GENERAUTO="CON"))');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Contrat');
      SetControlText('TAFF_AFFAIRE', 'Contrat');
    End
    else if Valeur = 'GRP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('XX_WHERE', ' AND AFF_AFFAIRE0 IN ("W","A")')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1)IN ("W","A")');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Affaire');
      SetControlText('TGP_AFFAIRE', 'Affaire');
    end
    else if Valeur = 'AFF' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'A')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'A');
      SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("A", "")');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Chantier');
      SetControlText('TAFF_AFFAIRE', 'Chantier');
    end
    else if Valeur = 'PRO' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'P')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'P');
      SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1)="P"');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Appel d''offre');
      SetControlText('TAFF_AFFAIRE', 'Appel d''offre');
    end
    else
    Begin
      SetControlText('XX_WHERE', '');
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', '')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', '');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Affaire');
      SetControlText('TAFF_AFFAIRE', 'Affaire');
    end;
  end;

end;



procedure TOF_BTSELFAC_MUL.OnClose ;
begin
  TheOldFac.free;
  TOBFacturable.free;
  TOBmemoFact.free;
  TOBAFACTURER.free;
  Inherited ;
end ;

procedure TOF_BTSELFAC_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELFAC_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELFAC_MUL.SetEventsGridListe(Etat: boolean);
begin
  if Etat then
  begin
  	Fliste.OnDblClick := FlisteDblClick
  end else
  begin
  	Fliste.OnDblClick := nil;
  end;
end;

procedure TOF_BTSELFAC_MUL.GetControls;
begin
  Fliste := THDbGrid(GetCOntrol('Fliste'));
  GS := THgrid(GetCOntrol('GS'));
  Blance := TToolbarButton97 (GetControl('BOUVRIR'));
  GS.visible := false;
	MnzPiece := TmenuItem (GetControl('mnZPiece'));
  BEfface := TToolbarButton97 (GetControl('BEFFACEAFF1'));
  //
  if GetControl('GP_DOMAINE') IS THValComboBox then Domaine:=THValComboBox(GetControl('GP_DOMAINE'));
  if Domaine<>Nil then PositionneDomaineUser(Domaine) ;
  //
end;


procedure TOF_BTSELFAC_MUL.FlisteDblClick(Sender: TObject);
var Affaire,Naturepieceg,Souche,TypeFac,Sql : string;
    Numero,IndiceG : integer;
    QQ : TQuery;
    TOBFacture : TOB;
    DevisPrinc : string;
    MessAttente : TFsplashScreen;
    sortie : TResultS;
    CBtnTxt         : PCustBtnText;
    NatS,SoucheS: string;
    NumS,indiceS : Integer;
begin
	if not ControleArticleEcartFact then
  BEGIN
//    PgiError (TraduireMemoire('L''article d''écart est invalide ou non renseigné#13#10Veuillez le définir'),Traduirememoire('Gestion d''écart'));
    exit;
  END;
  if ModeTraitement = TmtCloture then
  begin
    if not ConfirmationCloture then exit;
    if not SaisieDateCloture (DateCloture) then exit;
  end;
  //
  TheOldFac.clearDetail;
  TOBFacturable.clearDetail;
  //
  DevisPrinc := '';
  TOBFacture := TOB.Create('LA MEMO FACTURE',nil,-1);

  With TFMul(Ecran) do
  begin
    {$IFDEF EAGLCLIENT}
    Q.TQ.Seek(FListe.Row-1) ;
    {$ENDIF}
    TypeFac := Q.findField('AFF_GENERAUTO').AsString;
    Affaire := Q.FindField('GP_AFFAIRE').asstring;
    NaturepieceG := Q.FindField('GP_NATUREPIECEG').asstring ;
    Souche := Q.FindField('GP_SOUCHE').asstring ;
    Numero := Q.FindField('GP_NUMERO').asInteger ;
    IndiceG := Q.FindField('GP_INDICEG').asInteger ;
  end;
  fTypeFacturation := TypeFac;
  if (Pos(TypeFac,'AVA;DAC;')>0) then
  begin
    Avancement := True;
  end else
  begin
    Avancement := false;
  end;

  if (not ReajusteDossierfac) and (GetParamSocSecur('SO_BTONLYONESITU',false)) then
  begin
    if ExisteSituationProv (Avancement,NaturePieceg,Souche,Numero,IndiceG,NatS,SoucheS,NumS,indiceS) then
    begin
      system.New(CBtnTxt);
      TRY
        CBtnTxt[mbCust1] := 'Abandonner la saisie';
        CBtnTxt[mbCust3] := 'Récupérer la saisie précédente';
        Case ExMessageDlg('Une facture provisoire est déjà en attente.Que désirez-vous faire ?' + #13#10, mtxWarning, [mbCust1,mbCust3], 0, mbCust2, Application.Icon, CBtnTxt) Of
          mrCust1: exit;
          mrCust3: RecupSaisie := True;
        else
          exit;
        End;
      FINALLY
        Dispose(CBtnTxt);
      END;
    end;
  end;
  if AcompteNonRegle (Naturepieceg,Souche,numero,indiceg) then
  begin
      system.New(CBtnTxt);
      TRY
        CBtnTxt[mbCust1] := 'Abandonner la saisie';
        CBtnTxt[mbCust3] := 'Continuer';
        Case ExMessageDlg('L''acompte demandé n''a pas été encore réglé.Que désirez-vous faire ?' + #13#10, mtxWarning, [mbCust1,mbCust3], 0, mbCust2, Application.Icon, CBtnTxt) Of
          mrCust1: exit;
          mrCust3: ;
        else
          exit;
        End;
      FINALLY
        Dispose(CBtnTxt);
      END;
  end;
  if not RecupSaisie then
  begin
  RefPieceSel := EncoderefSel (NaturePieceg,Souche,Numero,IndiceG);

  QQ := OpenSql ('SELECT BMF_DEVISPRINC FROM BTMEMOFACTURE WHERE BMF_DEVIS="'+RefPieceSel+'"',true);
  if not QQ.eof then
  begin
    // on change la reference de la piece initiale par le devis initial de la facture
    DevisPrinc  := QQ.findfield('BMF_DEVISPRINC').AsString;
  end;
  ferme(QQ);
  // Phase 1 -- On recup les devis d'une facture existante
  if DevisPrinc <> '' then
  begin
    Sql := 'SELECT * FROM BTMEMOFACTURE WHERE BMF_DEVISPRINC="'+DevisPrinc+'"';
    TOBFacture.LoadDetailDBFromSQL ('BTMEMOFACTURE',Sql,false);
    TRY
      RefPieceSel := DevisPrinc;
      if not Chargefacturation(TOBFacture) then
      begin
      	PgiInfo ('Facturation impossible. La facturation à été cloturé',ecran.Caption);
      	exit;
      end;
    FINALLY
      TOBFacture.free;
    END;
  end else
  begin
    Sortie := ExisteFacturationOld (Affaire,Naturepieceg,Souche,Numero,IndiceG);
    if Sortie = Btrsok then
    begin
      if ControleEtatDevisPrinc (NaturepieceG,Souche,Numero,IndiceG) then
      begin
        MessAttente := TFsplashScreen.Create (ecran);
        MessAttente.Label1.Caption := 'Récupération de la structure de la facture en cours ...';
        MessAttente.Animate1.Active := true;
        MessAttente.Show;
        MessAttente.Refresh;
        DevisPrinc := EncoderefSel(NaturepieceG,Souche,Numero,IndiceG);
        Sql := 'SELECT ##TOP 1## BMF_DEVISPRINC FROM BTMEMOFACTURE WHERE BMF_DEVISPRINC="'+DevisPrinc+'"';
        if ExisteSQL(Sql) then
        begin
        	Sql := 'SELECT * FROM BTMEMOFACTURE WHERE BMF_DEVISPRINC="'+DevisPrinc+'"';
          TOBFacture.LoadDetailDBFromSQL ('BTMEMOFACTURE',Sql,false);
          TRY
            RefPieceSel := DevisPrinc;
            if not Chargefacturation(TOBFacture) then
            begin
              PgiInfo ('Facturation impossible. La facturation à été cloturé',ecran.Caption);
              exit;
            end;
          FINALLY
            TOBFacture.free;
          END;
        end else
        begin
        	ChargeOldfacturation (NaturePieceG,Souche,Numero,IndiceG);
        end;
        MessAttente.free;
      end else
      begin
      	Exit;
      end;
    end else if Sortie = BtrsNone then
    begin
        if ReajusteDossierfac then Exit;
      NewFacturation (NaturepieceG,Souche,Numero,IndiceG);
    end else if Sortie = BtrsAnnul then
    begin
      exit;
    end;
  end;
  TheOldFac.clearDetail;
  // Phase 2 -- Selection des devis non facturés
    if (not ReajusteDossierfac) and (not RecupSaisie) then
    begin
  ChargeLesDevisNonFacture (Affaire,TypeFac,Numero,IndiceG);
    end;
  beforeSelection;
  // Now c'est correct -- On peut afficher la grille de multiselection
  AfficheLagridGS;
  GS.Visible := true;
  end else
  begin
    // recupération de la situation précédente
    Definifacturation(NatS,SoucheS,NumS,indiceS);
    BlanceClick(self);
  end;
//  if TOBfacturable.detail.count = 1 then BlanceClick (self);

end;

procedure TOF_BTSELFAC_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
  Aff0 := THEdit(GetControl('AFFAIRE0'));
  Aff := THEdit(GetControl('GP_AFFAIRE'));
  Aff1 := THEdit(GetControl('GP_AFFAIRE1'));
  Aff2 := THEdit(GetControl('GP_AFFAIRE2'));
  Aff3 := THEdit(GetControl('GP_AFFAIRE3'));
  Aff4 := THEdit(GetControl('GP_AVENANT'));

end;

function TOF_BTSELFAC_MUL.Chargefacturation(TOBFacture: TOB) : boolean;
var Sql : string;
    TOBDevis,TOBL : TOB;
    Indice : integer;
    QQ : TQuery;
begin
	result := true;
  For Indice := 0 to TOBFacture.detail.count -1 do
  begin
    TOBL := TOBFacture.detail[Indice];
    Sql := Sqlbase+' WHERE '+WhereRefPieceSel(TOBL);
    QQ := OPenSql (Sql,True);
    if not QQ.eof then
    begin
      TOBDevis := TOB.Create ('PIECE',nil,-1);
      TOBdevis.AddChampSupValeur ('AFF_GENERAUTO','');
      TOBDevis.selectDb('',QQ);
      if (Indice = 0) and (TOBDevis.GetValue('AFF_ETATAFFAIRE')='TER') then
      begin
      	TOBDevis.free;
      	ferme(QQ);
      	Result := false;
        exit;
      end;
      TOBDevis.changeparent(TOBFacturable,-1);
      TOBDevis.AddChampSupValeur ('SELECTED','X');
      if indice = 0 then TOBReference := TOBdevis;
    end;
    ferme(QQ);
  end;
end;

function TOF_BTSELFAC_MUL.WhereRefPieceSel(TOBL: TOB): string;
var cledoc : r_cledoc;
begin
   DecodeRefPiece(TOBL.GetValue('BMF_DEVIS'),Cledoc);
   result := WherePiece (cledoc,ttdpiece,true);
end;

procedure TOF_BTSELFAC_MUL.ChargeLesDevisNonFacture (Affaire,TypeFac: string; NumSelected,Indice : integer);
var
    Part0,Part1,Part2,Part3,Avenant : string;
    Sql : string;
    TOBLocal : TOB;
begin
  BTPCodeAffaireDecoupe(Affaire,Part0,Part1,Part2,Part3,Avenant,taModif,false);
  if Avenant = '00' then
  begin
    TOBLOcal := TOB.Create('LES DOCUMENTS',nil,-1);
    Sql := SqlBase + ' WHERE GP_AFFAIRE1="'+Part1+'" AND GP_AFFAIRE2="'+Part2+'" AND GP_AFFAIRE3="'+Part3+'" AND '+
          'GP_NATUREPIECEG = "DBT" AND AFF_ETATAFFAIRE="ACP" '+
          'AND NOT EXISTS (SELECT BMF_DEVIS FROM BTMEMOFACTURE WHERE BMF_NATUREPIECEG=GP_NATUREPIECEG AND '+
                           'BMF_SOUCHE=GP_SOUCHE AND BMF_NUMERO=GP_NUMERO AND BMF_INDICEG=GP_INDICEG) '+
          'AND (GP_NUMERO <> '+IntToStr(Numselected)+' OR GP_INDICEG <> '+IntToStr(Indice)+') '+
          'AND AFF_GENERAUTO="'+TypeFac+'" '+
          'AND AFF_ADMINISTRATIF <> "X" AND AFF_MODELE <> "X"';
    TOBLOcal.LoadDetailDBFromSQL ('PIECE',Sql,true);
    if TOBlocal.detail.count > 0 then
    begin
      Indice := 0;
      repeat
        if not ExistDevis (TOBlocal.detail[Indice]) then
        begin
          TOBLOcal.detail[Indice].AddChampSupValeur ('SELECTED','-');
          TOBLocal.detail[Indice].changeparent (TOBfacturable,-1);
        end else inc(indice);
      until Indice >= TOBLocal.detail.count;
    end;
    TOBLocal.free;
  end;
end;

procedure TOF_BTSELFAC_MUL.NewFacturation(Naturepieceg, Souche: string;Numero, Indice: integer);
var Sql : string;
    Piece : string;
    cledoc : r_cledoc;
    TOBDevis : TOB;
    QQ : Tquery;
begin
   Piece := EncoderefSel (Naturepieceg,Souche,Numero,Indice);
   DecodeRefPiece(Piece,Cledoc);
   Sql := SqlBase + ' WHERE '+WherePiece (cledoc,ttdpiece,true);
   QQ := OpenSql (Sql,true);
   if not QQ.eof then
   begin
      TOBDevis := TOB.Create ('PIECE',nil,-1);
      TOBDevis.selectDb('',QQ);
      TOBDevis.changeparent(TOBFacturable,-1);
      TOBdevis.AddChampSupValeur ('SELECTED','X');
      TOBReference := TOBdevis;
   end;
   Ferme (QQ);
end;

procedure TOF_BTSELFAC_MUL.InitGrilleGS;
var Col,Dec, FixedWidth, Larg : integer ;
    tal : TAlignment ;
    STitre,St,Ch,stA,FF,Typ,FPerso : string ;
    NomList,FRecordSource,FLien,FSortBy,FFieldList,FLargeur,FAlignement,FParams : string ;
    Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul,OkTri,OkNumCol : boolean ;
    Nom : string;
    lescolonnes : string;
    Indice : integer;
    Ftitre,tt,NC : Hstring;
begin
  FixedWidth:=6;
  GS.ColCount:=60;
  Col:=1; STitre:=' '; ColName:='Fixed';

  With TFMul(Ecran) do
  BEGIN
    NomList:=Q.Liste;
    ChargeHListe(NomList,FRecordSource,FLien,FSortBy,FFieldList,FTitre,FLargeur,FAlignement,FParams,tt,NC,FPerso,OkTri,OkNumCol);
    GS.Titres.Add(FFieldList) ;
    While Ftitre<> '' do
    BEGIN
      StA:=ReadTokenSt(FAlignement);
      St:=ReadTokenSt(Ftitre);
      Ch:=ReadTokenSt(FFieldList);
      Larg:=ReadTokenI(FLargeur);
      tal:=TransAlign(StA,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
      if OkVisu then
      BEGIN
        GS.Cells[Col,0]:=St ;
        GS.ColAligns[Col]:=tal;
        GS.ColWidths[Col]:=Larg*GS.Canvas.TextWidth('W') ;
        if OkLib then GS.ColFormats[Col]:='CB=' + Get_Join(Ch);
        Typ:=ChampToType(Ch) ;
        if (Typ='INTEGER') or (Typ='SMALLINT') or (Typ='DOUBLE') then GS.ColFormats[Col]:=FF ;
        STitre:=STitre+';'+St;
        ColName:=ColName+';'+Ch;
        inc (Col);
      END;
    END;
    GS.ColCount:=Col ;
    GS.ColWidths[0]:=FixedWidth;
  END ;
  //
  lescolonnes := ColName;
  Nom := READTOKENST (LesColonnes);
  Indice := 0;
  repeat
    TraiteChampAffaireGrid (GS,Nom,indice);
    Nom := READTOKENST (LesColonnes);
    inc(Indice);
  until nom='';
  GS.Options:=GS.Options-[GoEditing,GoTabs,GoAlwaysShowEditor] ;
  GS.Options:=GS.Options+[GoRowSelect] ;
  GS.MultiSelect := true;
  GS.TwoColors := true;
end;

procedure TOF_BTSELFAC_MUL.AfficheLagridGS;
var Indice : integer;
begin
  SetEventsGridGS (false);
  //
  GS.ClearSelected;
  GS.VidePile (false);
  GS.rowCount := TOBFacturable.detail.count+1;
  For Indice := 0 to TOBFacturable.detail.count -1 do
  begin
    TOBFacturable.detail[Indice].PutLigneGrid (GS,Indice+1,false,false,ColName);
    GS.Objects [0,Indice+1] := TOBFacturable.detail[Indice];
    if TOBFacturable.detail[Indice].GetValue('SELECTED')='X' then GS.FlipSelection (Indice+1);
  end;
  //
  SetEventsGridGS (True);
end;

Procedure TOF_BTSELFAC_MUL.TraiteChampAffaireGrid ( Gr : THGrid; Nom : string; Col : integer);
Var  Lg,ind,NumPart : integer;
     NomChamp,Libelle : string;
     Visible : Boolean;
begin
if nom = '' then Exit;
lg := Length(Nom); ind:= Pos ('_',Nom);
if (lg =0) or (ind = 0) then  exit;
Nomchamp := Copy (Nom, ind+1,lg-ind);
Visible:=False; //mcd 14/02/03
If ((Nomchamp = 'AFFAIRE1') Or (Nomchamp = 'AFFAIRE2') Or (Nomchamp = 'AFFAIRE3')) Then
  Begin
  NumPart := StrToInt(Copy(NomChamp, 8,1));
  Case NumPart Of
      1 : Begin Visible:=True; Libelle:=VH_GC.CleAffaire.Co1Lib;  End;
      2 : Begin Visible:=VH_GC.CleAffaire.Co2Visible; Libelle:=VH_GC.CleAffaire.Co2Lib; End;
      3 : Begin Visible:=VH_GC.CleAffaire.Co3Visible; Libelle:=VH_GC.CleAffaire.Co3Lib; End;
      End;
  if not(Visible) then Gr.colwidths[Col]:=0;
  Gr.Cells[Col,0]:=Libelle;
  End;
end;


procedure TOF_BTSELFAC_MUL.SetEventsGridGS(etat: boolean);
begin
  if Etat  then
  begin
    GS.OnBeforeFlip := FlipThisRow;
    GS.PostDrawCell := GSPostDrawCell;
    GS.OnRowEnter   := GSRowEnter;
    GS.OnRowExit    := GSRowExit;
  end else
  begin
    GS.OnBeforeFlip := nil;
    GS.PostDrawCell := nil;
    GS.OnRowEnter   := nil;
    GS.OnRowExit    := nil;
  end;
end;

function TOF_BTSELFAC_MUL.IsPieceSelectionnable(Arow : integer) : boolean;
var TOBL : TOB;
begin
  result := true;
  TOBL := TOB(GS.Objects [0,Arow]);
  if TOBreference <> nil then
  begin
    if TOBreference.getValue('PRG_TYPERG')<> TOBL.getValue('PRG_TYPERG') then
    begin
      PGIInfo('Le mode de retenue de garantie de ce document est différent de celui du devis de base');
      result := false;
      Exit;
    end;

    if TOBreference.getValue('RUPTMILLIEME')<> TOBL.getValue('RUPTMILLIEME') then
    begin
      PGIBox('Erreur de Millième');
      result := false;
      Exit;
    end;
  end;
  if Existefac (TOBL) then
  BEGIN
    PgiBox('Impossible. Il existe une facturation sur ce devis');
    result := false;
  end;
end;

procedure TOF_BTSELFAC_MUL.FlipThisRow(Sender: TObject; ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
begin
  if ReajusteDossierfac then
  begin
    cancel := True;
    Exit;
  end;
  if GS.IsSelected (Arow) then
  begin
    // Cas ou l'on veut enlever un devis de la sélection
    // On verifie alors s'il existe des lignes de LIGNEFAC pour ce devis
    TOBL := TOB(GS.Objects [0,Arow]);
    if Existefac (TOBL) then
    BEGIN
      PgiBox('Impossible. Il existe une facturation sur ce devis');
      cancel := true;
    end;
    if TOBL = TOBreference then TOBreference := FindFirstSelected;
  end else
  begin
    If not IsPieceSelectionnable(Arow) then begin cancel := true;exit; end;
    if TOBreference = nil then TOBreference := TOB(GS.Objects [0,Arow]);
  end;
end;

function TOF_BTSELFAC_MUL.FindFirstSelected : TOB;
var Indice : integer;
begin
  result := nil;
  for Indice := 1 to GS.rowCount do
  begin
    if GS.IsSelected (Indice) then
    begin
      result := TOB(GS.Objects [0,Indice]);
      break;
    end;
  end;
end;

function TOF_BTSELFAC_MUL.Existefac(TOBL: TOB): boolean;
var req : string;
begin
	// Sur Affaire type avancement
  result := false;
  req := 'SELECT ##TOP 1## BLF_NUMORDRE FROM LIGNEFAC WHERE BLF_NATUREPIECEG="'+
         TOBL.GetValue('GP_NATUREPIECEG')+'" AND BLF_SOUCHE="'+TOBL.GetValue('GP_SOUCHE')+'" '+
         'AND BLF_NUMERO='+InttoStr(TOBL.GetValue('GP_NUMERO'))+' AND BLF_INDICEG='+InttoStr(TOBL.GetValue('GP_INDICEG'));
  if ExisteSql(Req) then BEGIN result := True; Exit; END;
  req := 'SELECT ##TOP 1## BST_NUMEROSIT FROM BSITUATIONS WHERE BST_SSAFFAIRE="'+ TOBL.GetValue('GP_AFFAIREDEVIS')+'"';
  if ExisteSql(Req) then BEGIN result := True; Exit; END;
  // Sur affaire type directe ou type avancement (antériorité V7)
  req := 'SELECT GL_PIECEPRECEDENTE FROM LIGNE WHERE (GL_NATUREPIECEG IN ("FBT","DAC")) AND '+
  			 '(GL_PIECEPRECEDENTE LIKE "%;'+TOBL.GetValue('GP_NATUREPIECEG')+';'+
         TOBL.GetValue('GP_SOUCHE')+';'+IntToStr(TOBL.GetValue('GP_NUMERO'))+';'+
         InttoStr(TOBL.GetValue('GP_INDICEG'))+';%")';
  if ExisteSql(Req) then BEGIN result := True; Exit; END;
end;

procedure TOF_BTSELFAC_MUL.SetEvents;
begin
  Blance.OnClick := BlanceClick;
  mnZPiece.OnClick := VoirPiece;
  BEfface.OnClick := BEffaceClick;
end;

procedure TOF_BTSELFAC_MUL.BlanceClick(Sender: Tobject);
begin
  if (RecupSaisie) then
  begin
    // Récupération d'une saisie
    TheTOB := TOBAFACTURER;
    AglLanceFiche ('BTP','BTSAISIEAVANC','','','MODIFSITUATION');
    TheTOB := Nil;
  end else if (GS.visible)  then
  begin
  	if not MontantDevZero then
    begin
      MiseANiveauDesDocuments;
    	EnregistreSelection;
    end else
    begin
      PgiBox('Certains documents dont le montant est nul doivent être recalculés...');
    end;
  end else
  begin
    PgiBox('Vous n''avez pas encore sélectionné de devis');
  end;
end;

procedure TOF_BTSELFAC_MUL.EnregistreSelection;
var Arow,Numordre : integer;
    TOBP : TOB;
    io: TIoErr;
    DocOk : boolean;
    TRAITEMENT,MessageErreur : String;
begin
  Numordre := 0;
  TOBmemoFact.clearDetail;
  TOBAfacturer.clearDetail;
  DocOk := true;
  for Arow := 1 to GS.rowCount do
  begin
    if GS.IsSelected (Arow) then
    begin
      Inc(NumOrdre);
      TOBP := TOB(GS.objects[0,Arow]);
      DocOk := AddSelection(TOBP,NumOrdre,MessageErreur);
      if not DocOk then break;
    end;
  end;
  if not DocOk then
  BEGIN
  	PgiInfo (MessageErreur);
    BchercheClick;
    exit;
  END;
  if not ReajusteDossierfac then
  begin
  io := TRANSACTIONS (Majmemorisation,0);
  if io = OeOk then
  begin
    if TOBAFacturer.detail.count > 0 then
    begin
      // Go la saisie ..go
      TheTOB := TOBAFACTURER;
      if ModeTraitement = TmtCloture then TRAITEMENT :='CLOTURE' else TRAITEMENT :='';
      AglLanceFiche ('BTP','BTSAISIEAVANC','','',TRAITEMENT);
      TheTOB := Nil;
    end;
    BchercheClick;
  end else
  begin
    PGIError ('ATTENTION : Sélection non enregistrée');
    BchercheClick;
  end;
  end else
  begin
    // reajustement dossier de facturation
    TheTOB := TOBAFACTURER;
    AglLanceFiche ('BTP','BTSAISIEAVANC','','','REAJUSTEDOSSIER');
    TheTOB := Nil;
  end;
end;

function TOF_BTSELFAC_MUL.AddSelection(TOBP : TOB; NumOrdre : integer; var MessageErreur : string) : boolean;
var TOBMF : TOB;
    Document : string;
    QQ : TQuery;
    piece : string;
    cledoc : r_cledoc;
begin
	result := true;
  // Constitution de la Memorisation
  TOBMF := TOB.Create ('BTMEMOFACTURE',TOBmemoFact,-1);
  TOBMF.putValue('BMF_AFFAIRE',TOBP.GetValue('GP_AFFAIRE'));
  TOBMF.putValue('BMF_AFFAIRE1',TOBP.GetValue('GP_AFFAIRE1'));
  TOBMF.putValue('BMF_AFFAIRE2',TOBP.GetValue('GP_AFFAIRE2'));
  TOBMF.putValue('BMF_AFFAIRE3',TOBP.GetValue('GP_AFFAIRE3'));
  TOBMF.putValue('BMF_DEVISPRINC',RefPieceSel);
  Document := EncoderefSel (TOBP);
  TOBMF.putValue('BMF_DEVIS',document);
  TOBMF.putValue('BMF_NATUREPIECEG',TOBP.GetValue('GP_NATUREPIECEG'));
  TOBMF.putValue('BMF_SOUCHE',TOBP.GetValue('GP_SOUCHE'));
  TOBMF.putValue('BMF_NUMERO',TOBP.GetValue('GP_NUMERO'));
  TOBMF.putValue('BMF_INDICEG',TOBP.GetValue('GP_INDICEG'));
  TOBMF.putValue('BMF_NUMORDRE',NumOrdre);
  // constitution de la TOB des pièces a facturer
  TOBMF := TOB.create ('PIECE',TOBAFACTURER,-1);
  Piece := EncodeRefPiece (TOBP);
  DecodeRefPiece (Piece,Cledoc);
  QQ := OpenSql ('SELECT PIECE.*,AFF_GENERAUTO,AFF_OKSIZERO,AFF_MANDATAIRE FROM PIECE '+
  							 'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIREDEVIS '+
                 'WHERE '+WherePiece (Cledoc,ttdPiece,false),true);
  if not QQ.eof then
  begin
    TOBMF.SelectDB ('',QQ);
    TOBAfacturer.AddChampSupValeur ('GP_AFFAIRE',TOBMF.GetValue('GP_AFFAIRE'));
    TOBAfacturer.AddChampSupValeur ('GP_TIERS',TOBMF.GetValue('GP_TIERS'));
    TOBAfacturer.AddChampSupValeur ('LIBELLEAFFAIRE',GetChampsAffaire(TOBMF.GetValue('GP_AFFAIRE'),'AFF_LIBELLE'));
    TOBAfacturer.AddChampSupValeur ('LIBELLETIERS',RecupLibelleTiers(TOBMF.GetValue('GP_TIERS')));
    TOBAfacturer.AddChampSupValeur ('AFF_GENERAUTO',TOBMF.GetValue('AFF_GENERAUTO'));
    TOBAfacturer.AddChampSupValeur ('DATECLOTURE',DateCloture);
    TOBAfacturer.AddChampSupValeur ('AFF_MANDATAIRE',TOBMF.GetValue('AFF_MANDATAIRE'));
  end;
  ferme(QQ);
end;


procedure TOF_BTSELFAC_MUL.Majmemorisation;
begin
  SupprimeMemo ;
  if TOBmemofact.detail.count > 0 then if not TOBmemofact.InsertDB (nil) then V_PGI.IOError := OeUnknown;
end;

procedure TOF_BTSELFAC_MUL.SupprimeMemo;
begin
  V_PGI.ioerror := OeOk;
  ExecuteSql ('DELETE FROM BTMEMOFACTURE WHERE BMF_DEVISPRINC="'+RefPieceSel+'"');
end;

procedure TOF_BTSELFAC_MUL.BchercheClick;
begin
  TToolBarButton97(GetCOntrol('BCHERCHE')).Click;
end;

procedure TOF_BTSELFAC_MUL.beforeSelection;
var Indice : integer;
    TOBL,TOBMilliemes : TOB;
    QQ : TQuery;
    cledoc : r_cledoc;
begin
  TOBMilliemes := TOB.Create ('LES MILLIEMEES',nil,-1);
  for Indice := 0 to TOBFacturable.detail.count -1 do
  begin
    TOBL := TOBfacturable.detail[Indice];
 		TOBL.AddChampSupValeur('AVANCSAISIE',0);
 		TOBL.AddChampSupValeur('AVANCPREC',0);
 		TOBL.AddChampSupValeur('RUPTMILLIEME','');
    //
    TOBMilliemes.ClearDetail;
    Cledoc := TOB2CleDoc (TOBL);
    QQ := OPenSql ('SELECT BPM_CATEGORIETAXE,BPM_FAMILLETAXE,BPM_MILLIEME FROM BTPIECEMILIEME WHERE '+WherePiece (Cledoc,TTdRepartmill,false),True);
    if not QQ.eof then
    begin
     TOBMilliemes.LoadDetailDB ('BTPIECEMILIEME','','',QQ,false);
     TOBL.putValue('RUPTMILLIEME',EncodeRepartTva (TOBMilliemes));
    end;
    Ferme(QQ);
  end;
  TOBMilliemes.free;
end;

function TOF_BTSELFAC_MUL.IsPresentDansFactureOld (Var TheRefDevis : string; cledoc : r_cledoc) : boolean;
var SQl,ClePiece : String;
		QQ : TQuery;
    TOBDFact : TOB;
begin
	result := false;
	TOBDfact := TOB.create ('LES FACTURES', nil,-1);
  TRY
    ClePiece :=';'+Cledoc.NaturePiece+';'+Cledoc.souche+';'+IntToStr(Cledoc.NumeroPiece)+';';
    Sql := 'SELECT DISTINCT BST_SSAFFAIRE FROM LIGNE '+
    			 'LEFT JOIN BSITUATIONS ON BST_NATUREPIECE=GL_NATUREPIECEG AND '+
           'BST_SOUCHE=GL_SOUCHE AND BST_NUMEROFAC=GL_NUMERO WHERE '+
           'GL_NATUREPIECEG IN ("FBT","FBP") AND GL_PIECEPRECEDENTE LIKE "%'+ClePiece+'%"';
    QQ := OpenSql (Sql,true,-1,'',True);
    if not QQ.eof then
    begin
      TOBDfact.LoadDetailDB('UNE LIGNE','','',QQ,false);
      TheRefDevis := TOBDfact.detail[0].getValue('BST_SSAFFAIRE');
      result := true;
    end;
    ferme (QQ);
  FINALLY
  	TOBDFact.free;
  END;
end;

function TOF_BTSELFAC_MUL.ExisteFacturationOld(var Affaire,Naturepieceg,Souche: string; var Numero, Indice: integer): TResultS;
var QQ : TQuery;
    TheRefDevis : string;
    TOBTMP,TOBSEL : TOB;
    req : string;
    cledoc : r_cledoc;
begin
  result := BtrsNone;
  TOBTMP := TOB.Create ('LES SITUATIONS SUR LAFFAIRE',nil,-1);
  TheRefDevis := '';
  QQ := OpenSql ('SELECT GP_AFFAIREDEVIS FROM PIECE WHERE GP_NATUREPIECEG="'+NaturepIeceg+'" AND '+
                 'GP_SOUCHE="'+SOuche+'" AND GP_NUMERO='+InttoStr(Numero)+' AND GP_INDICEG='+InttOstr(Indice),True,-1,'',true);
  if not QQ.eof then
  begin
    TheRefDevis := QQ.findField('GP_AFFAIREDEVIS').AsString;
  end;
  ferme (QQ);

  if not avancement then
  begin
  	result := BtrsNone;
    exit;
  end;

  if TheRefDevis <> '' then
  begin
    QQ := OpenSQL ('SELECT * FROM BSITUATIONS WHERE '+
                   'BST_SSAFFAIRE="'+TheRefDevis+'" AND '+
                   'BST_VIVANTE="X" '+
                   'ORDER BY BST_NUMEROFAC',true,-1,'',true);
    if not QQ.eof then
    begin
      TheOldFac.LoadDetailDB ('BSITUATIONS','','',QQ,false);
      result := BtrsOK;
    end;
    ferme (QQ);
  end;
  // cas de figure ou il n'existe pas de facturation précédente avce le devis indiqué comme devis principal
  if result <> Btrsok then
  begin
    if ReajusteDossierfac then
    begin
  	  result := BtrsNone;
      Exit;
    end;
  	//
    cledoc.NaturePiece := Naturepieceg;
    cledoc.Souche      := Souche;
    cledoc.NumeroPiece := Numero;
    cledoc.Indice := Indice;
    if IsPresentDansFactureOld (TheRefDevis,cledoc) then
    begin
      Naturepieceg := Copy(TheRefDevis,2,3);
      Souche := Copy(TheRefDevis,5,3);
      Numero := Strtoint(Copy(TheRefDevis,8,8));
      Indice := 0;
      result := BtrsOK;
      QQ := OpenSQL ('SELECT * FROM BSITUATIONS WHERE '+
                     'BST_SSAFFAIRE="'+TheRefDevis+'" '+
                     'BST_VIVANTE="X" '+
                     'ORDER BY BST_NUMEROFAC',true,-1,'',true);
      if not QQ.eof then
      begin
        TheOldFac.LoadDetailDB ('BSITUATIONS','','',QQ,false);
        result := BtrsOK;
      end;
      ferme (QQ);
      exit;
    end;
    //
    REq := 'SELECT B1.* FROM BSITUATIONS B1 WHERE '+
           'B1.BST_AFFAIRE="'+Affaire+'" AND B1.BST_VIVANTE="X" AND '+
           'B1.BST_NUMEROSIT=('+
            'SELECT MAX(B2.BST_NUMEROSIT) FROM BSITUATIONS B2 '+
            'WHERE B2.BST_SSAFFAIRE=B1.BST_SSAFFAIRE AND B2.BST_VIVANTE="X"'+
            ') '+
            'ORDER BY B1.BST_DATESIT';
    QQ := OpenSQL (req,true,-1,'',true);
    if not QQ.eof then
    begin
      TOBTMP.loadDetailDb ('SITUATIONS','','',QQ,false);
    end;
    ferme (QQ);
    // TODO (evacuer ici les factures sur les devis termine ...cf cloture)
    VerifDevisFromFacturation (TOBTMP);
    //
    if TOBTMP.detail.count <> 0 then
    begin
      TOBSEL := ChoixSelectionFacture (TOBTMP);
      if TOBSel <> nil then
      begin
        TheRefDevis := TOBSel.getValue('BST_SSAFFAIRE');
        if TheRefDevis <> '' then
        begin
          QQ := OpenSQL ('SELECT * FROM BSITUATIONS WHERE '+
                         'BST_SSAFFAIRE="'+TheRefDevis+'" AND '+
                         'BST_VIVANTE="X" '+
                         'ORDER BY BST_NUMEROFAC',true,-1,'',true);
          if not QQ.eof then
          begin
            TheOldFac.LoadDetailDB ('BSITUATIONS','','',QQ,false);
            Naturepieceg := Copy(TheRefDevis,2,3);
            Souche := Copy(TheRefDevis,5,3);
            Numero := Strtoint(Copy(TheRefDevis,8,8));
            Indice := 0;
            result := BtrsOK;
          end;
          ferme (QQ);
        end else
        begin
          TOBSel.free;
        end;
      end else
      begin
        result := BtrsAnnul;
      end;
    end;
  end;
end;

function TOF_BTSELFAC_MUL.ExistDevis (cledoc : r_cledoc): boolean;
begin
  result :=  (TOBFacturable.FindFirst(['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'],
                                [cledoc.NaturePiece,cledoc.Souche,cledoc.NumeroPiece,Cledoc.Indice],true)<>nil);
end;

procedure TOF_BTSELFAC_MUL.ChargeOldfacturation (Naturepieceg,Souche : string; var Numero,IndiceG : integer);
var TOBDevis,TOBFacture,TheLastFacture,TOBL : TOB;
    indice,IndFac : integer;
    QQ : TQuery;
    cledoc : R_CLEDOC;
    first : boolean;
begin
  first := true;
  TOBfacture := TOB.Create ('LA FACTURE',nil,-1);
  TRY
    for IndFac := 0 to TheOldFac.detail.count -1 do
    begin
      TOBFacture.clearDetail;
      TheLastFacture := TheOldFac.detail[IndFac];
      TOBfacture.LoadDetailDBFromSQL ('LIGNE','SELECT DISTINCT (GL_PIECEPRECEDENTE) FROM LIGNE '+
                                      'WHERE GL_NATUREPIECEG="'+TheLastFacture.getValue('BST_NATUREPIECE')+'" AND '+
                                      'GL_SOUCHE="'+TheLastFacture.getValue('BST_SOUCHE')+'" AND '+
                                      'GL_NUMERO='+IntToStr(TheLastFacture.getValue('BST_NUMEROFAC')) + ' AND '+
                                      'GL_PIECEPRECEDENTE <> ""',false);
// modif brl pour pb regroupement sur DAC ABH PACE (cas du devis principal sans ligne article facturée)
//'GL_NUMERO='+IntToStr(TheLastFacture.getValue('BST_NUMEROFAC'))+' AND GL_TYPELIGNE="ART"',false);
      if TOBfacture.detail.count <> 0 then
      begin
      	if first then
        begin
        // Afin d'eviter tout problème on charge la pièce du document de reference
          QQ := OpenSql (SqlBase +' WHERE GP_NATUREPIECEG="'+Naturepieceg+'" AND '+
                         'GP_SOUCHE="'+Souche+'" AND GP_NUMERO='+InttoStr(Numero)+ ' AND '+
                         'GP_INDICEG='+InttOStr(IndiceG),true,-1,'',true);
          if not QQ.eof then
          begin
            TOBDevis := TOB.Create ('PIECE',nil,-1);
            TOBDevis.selectDb('',QQ);
            TOBDevis.changeparent(TOBFacturable,-1);
            TOBdevis.AddChampSupValeur ('SELECTED','X');
            TOBReference := TOBdevis;
            first := false;
          end;
        end;
        //
        for Indice := 0 to TOBFacture.detail.count -1 do
        begin
          TOBL := TOBFacture.detail[Indice];
// modif brl pour pb regroupement sur DAC ABH PACE (cas du devis principal sans ligne article facturée)
//          if TOBL.getValue('GL_TYPELIGNE')<>'ART' then continue;
          if TOBL.getValue('GL_PIECEPRECEDENTE')='' then continue;
          DecodeRefPiece (TOBL.getValue('GL_PIECEPRECEDENTE'),cledoc);
          //
          if not ExistDevis (Cledoc) then
          begin
            QQ := OpenSql (SqlBase +'WHERE GP_NATUREPIECEG="'+Cledoc.NaturePiece+'" AND '+
                           'GP_SOUCHE="'+Cledoc.Souche+'" AND GP_NUMERO='+InttoStr(cledoc.NumeroPiece)+ ' AND '+
                           'GP_INDICEG='+InttOStr(Cledoc.Indice)+' AND AFF_ETATAFFAIRE="ACP"',true,-1,'',true);
            if not QQ.eof then
            begin
              TOBDevis := TOB.Create ('PIECE',nil,-1);
              TOBDevis.selectDb('',QQ);
              TOBDevis.changeparent(TOBFacturable,-1);
              TOBdevis.AddChampSupValeur ('SELECTED','X');
            end;
          end;
          ferme(QQ);
        end;
      end;
    end;
  FINALLY
    TOBfacture.free;
  END;
end;

function TOF_BTSELFAC_MUL.ExistDevis(TOBL: TOB): boolean;
var cledoc: r_cledoc;
    TheRef : string;
begin
  TheRef := EncodeRefPiece (TOBL);
  DecodeRefPiece (TheRef,cledoc);
  result := existDevis (cledoc);
end;


procedure TOF_BTSELFAC_MUL.VoirPiece(Sender: TObject);
var cledoc : r_cledoc;
		Laction : TactionFiche;
    QQ : TQuery;
begin
  if GS.visible then
  begin
    cledoc.NaturePiece  := TOBFacturable.detail[GS.row-1].getValue('GP_NATUREPIECEG');
    cledoc.Souche  := TOBFacturable.detail[GS.row-1].getValue('GP_SOUCHE');
    cledoc.NumeroPiece  := TOBFacturable.detail[GS.row-1].getValue('GP_NUMERO');
    cledoc.Indice  := TOBFacturable.detail[GS.row-1].getValue('GP_INDICEG');
  end else
  begin
    With TFMul(Ecran) do
    begin
      {$IFDEF EAGLCLIENT}
      Q.TQ.Seek(FListe.Row-1) ;
      {$ENDIF}
      cledoc.NaturePiece := Q.FindField('GP_NATUREPIECEG').asstring ;
      cledoc.Souche := Q.FindField('GP_SOUCHE').asstring ;
      cledoc.NumeroPiece := Q.FindField('GP_NUMERO').asInteger ;
      cledoc.Indice := Q.FindField('GP_INDICEG').asInteger ;
    end;
  end;
  //
  if JaiLeDroitNatureGCModif(CleDoc.NaturePiece) then LAction := taModif else LAction := taConsult;
  //
  SaisiePiece(CleDoc, LAction,'','','',false,false,false,true);
  if GS.visible then
  begin
  	if LAction = taModif then
    begin
      QQ := OpenSql ('SELECT * FROM PIECE WHERE '+WherePiece(Cledoc,ttdpiece,true),true,-1,'',true);
      if not QQ.eof then TOBFacturable.detail[GS.row-1].SelectDB('',QQ);
      ferme (QQ);
      TOBFacturable.detail[GS.row-1].PutLigneGrid (GS,Gs.row,false,false,ColName);
    end;
  end else
  begin
    BchercheClick;
  end;
end;

function TOF_BTSELFAC_MUL.MontantDevZero: boolean;
var Arow : integer;
		TOBP : TOB;
begin
	result := false;
  for Arow := 1 to GS.rowCount do
  begin
    if GS.IsSelected (Arow) then
    begin
      TOBP := TOB(GS.objects[0,Arow]);
      if TOBP.GetValue('GP_TOTALHT')=0 then
      begin
      	result := true;
        break;
      end;
    end;
  end;
end;

procedure TOF_BTSELFAC_MUL.GSPostDrawCell(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var ARect: TRect;
		Triangle: array[0..2] of TPoint;
begin
  if (gdFixed in AState) and (ACol = 0) then
  begin
    Arect := GS.CellRect(Acol, Arow);
    Canvas.Brush.Color := GS.FixedColor;
    Canvas.FillRect(ARect);
    if (ARow = GS.row) then
    begin
      Canvas.Brush.Color := clBlack;
      Canvas.Pen.Color := clBlack;
      Triangle[1].X := ARect.Right - 2;
      Triangle[1].Y := ((ARect.Top + ARect.Bottom) div 2);
      Triangle[0].X := Triangle[1].X - 5;
      Triangle[0].Y := Triangle[1].Y - 5;
      Triangle[2].X := Triangle[1].X - 5;
      Triangle[2].Y := Triangle[1].Y + 5;
      if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle);
    end;
  end;
end;

procedure TOF_BTSELFAC_MUL.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
	GS.InvalidateRow(Ou);
end;

procedure TOF_BTSELFAC_MUL.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
	GS.InvalidateRow(Ou);
end;

procedure TOF_BTSELFAC_MUL.VerifDevisFromFacturation(TOBTMP: TOB);
var Indice : integer;
		TOBLS : TOB;
    Nature,Souche,noDevis : string;
    QQ : Tquery;
begin
	if TOBTmp.detail.count = 0 then exit;
	Indice := 0;
	repeat
    TOBLS := TOBTMP.detail[Indice];
    Nature := Copy(TOBLS.getValue('BST_SSAFFAIRE'),2,3);
    Souche := Copy(TOBLS.getValue('BST_SSAFFAIRE'),5,3);
    NoDevis := Copy(TOBLS.getValue('BST_SSAFFAIRE'),8,8);
    QQ := OpenSql ('SELECT AFF_ETATAFFAIRE,GP_REFINTERNE FROM PIECE LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIREDEVIS '+
   								 'WHERE GP_NATUREPIECEG="'+Nature+'" AND GP_SOUCHE="'+Souche+'" AND '+
	                 'GP_NUMERO='+NoDevis+' AND GP_INDICEG=0',true,-1,'',true);
    if Not QQ.eof then
    begin
    	if QQ.findfield('AFF_ETATAFFAIRE').asString='TER' then TOBLS.free else Inc(Indice);
    end else
    begin
    	Inc(Indice);
    end;
    ferme (QQ);
  until Indice >= TOBTmp.detail.count;
end;

function TOF_BTSELFAC_MUL.ControleEtatDevisPrinc (NaturepieceG,Souche : string; Numero,IndiceG : integer) : boolean;
var QQ : Tquery;
		Sql : string;
begin
  result := true;
  Sql := 'SELECT AFF_ETATAFFAIRE FROM PIECE LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIREDEVIS WHERE '+
         'GP_NATUREPIECEG="'+NaturepieceG+'" AND GP_SOUCHE="'+Souche+'" AND GP_NUMERO='+InttoStr(Numero)+
         ' AND GP_INDICEG='+InttoStr(IndiceG);
  QQ := OPenSql (Sql,True);
  if not QQ.eof then
  begin
    if QQ.FindField('AFF_ETATAFFAIRE').asString = 'TER' then
    begin
    	PgiInfo ('Facturation impossible. La facturation à été cloturé',ecran.Caption);
      result := false;
    end;
  end;
  ferme (QQ);
end;

function TOF_BTSELFAC_MUL.ConfirmationCloture: boolean;
var UneTOB : TOB;
		Modefacturation : string;
begin
  if fTypeFacturation  ='AVA' then
  begin
  	ModeFacturation := 'AVANCEMENT';
  end else if fTypeFacturation ='DAC' then
  begin
  	ModeFacturation := 'MEMOIREFINAL';
  end else
  begin
  	ModeFacturation := 'DIRECTE';
  end;
  UneTOB := TOB.create ('UNE DEMANDE',nil,-1);
  UneTOB.AddChampSupValeur('CONFIRMATION','-');
  UneTOB.AddChampSupValeur('MODECLOTURE',ModeFacturation);
  TheTOB := UneTOB;
  AGLLanceFiche('BTP','BTCONFIRMECLOTURE','','','');
  result := (UneTOB.GetValue('CONFIRMATION')='X');
  TheTOB := nil;
  UneTOB.free;
end;


function TOF_BTSELFAC_MUL.ControleArticleEcartFact: boolean;
var ARtEcart : string;
		QArt : TQuery;
begin
	result := true;
  ArtEcart := GetParamsoc('SO_ECARTFACTURATION');
  if ArtEcart = '' then
  begin
    PgiError (Traduirememoire('Impossible : l''article d''écart n''a pas été paramétré'));
    result := false;
  end;
  Qart := opensql('Select GA_ARTICLE from ARTICLE Where GA_CODEARTICLE="' + ArtEcart + '"', true,-1, '', True);
  if Qart.eof then
  begin
    PgiError (Traduirememoire('Impossible : l''article d''écart est invalide'));
    result := false;
  end;
  ferme (Qart);
end;

procedure TOF_BTSELFAC_MUL.MiseANiveauDesDocuments;

	function ControleOuvragesUnique (TOBP : TOB) : boolean;
  var Sql : string;
  begin
    result := true;
    Sql := 'SELECT BLO_UNIQUEBLO FROM LIGNEOUV WHERE BLO_NATUREPIECEG="'+TOBP.GetValue('GP_NATUREPIECEG')+'" AND '+
           'BLO_SOUCHE="'+TOBP.GetValue('GP_SOUCHE')+'" AND BLO_NUMERO='+intToStr(TOBP.GetValue('GP_NUMERO'))+ ' AND '+
           'BLO_INDICEG='+IntToStr(TOBP.GetValue('GP_INDICEG'))+' AND BLO_UNIQUEBLO=0';
    if ExisteSQL(Sql) then
    begin
      result := false;
    end;
  end;

  procedure AddMajUniqueBLO (TOBP,TOBPieceMan : TOB);
  var TOBAA : TOB;
  		QQ : Tquery;
      StSql : string;
      cledoc : r_cledoc;
      RefPiece : string;
  begin
  	TOBAA := TOB.Create ('PIECE',TOBPieceMan,-1);
    TOBAA.Dupliquer(TOBP,false,true);
    RefPiece := EncodeRefPiece(TOBP);
    DecodeRefPiece(RefPiece,cledoc );
    StSQL := 'SELECT * FROM LIGNEOUV WHERE ' + WherePiece (cledoc,ttdOuvrage,true);
    StSQL := StSQl + 'ORDER BY BLO_NUMLIGNE, BLO_N1, BLO_N2, BLO_N3, BLO_N4, BLO_N5';
    QQ := OpenSql(StSQL, True) ;
    TOBAA.LoadDetailDB  ('LIGNEOUV','','',QQ,false);
    Ferme (QQ);
  end;

var Arow : integer;
		TOBP,TOBPiecesMan : TOB;
begin
  TOBPiecesMan := TOB.Create ('LES PIECE MAN',nil,-1);
  TRY
    for Arow := 1 to GS.rowCount do
    begin
      if GS.IsSelected (Arow) then
      begin
        TOBP := TOB(GS.objects[0,Arow]);
        if not ControleOuvragesUnique (TOBP) then AddMajUniqueBLO (TOBP,TOBPiecesMan);
      end;
    end;
    if TOBPiecesMan.detail.count > 0 then
    begin
      TraiteManPieces (TOBPiecesMan,true);
    end;
  FINALLY
    TOBPiecesMan.free;
  END;
end;

procedure TOF_BTSELFAC_MUL.BeffaceClick(Sender: Tobject);
begin

  SetControlText('GP_TIERS', '');
  SetControlText('GP_AFFAIRE', '');
  SetControlText('GP_AFFAIRE1', '');
  SetControlText('GP_AFFAIRE2', '');
  SetControlText('GP_AFFAIRE2', '');
  SetControlText('GP_AVENANT', '');

end;


procedure TOF_BTSELFAC_MUL.Definifacturation(NaturePieceG,SoucheG: string; Numero,IndiceG : integer);
var QQ : TQuery;
    Nature,Souche,SQl : string;
    Num,indice : Integer;
    TOBMF : TOB;
begin
  TOBAFACTURER.ClearDetail;
  if Avancement then
  begin
    SQL := 'SELECT BST_NATUREPIECE,BST_SOUCHE,BST_NUMEROFAC FROM BSITUATIONS WHERE BST_NATUREPIECE="FBP" AND '+
                  'BST_SSAFFAIRE=('+
                  'SELECT GP_AFFAIREDEVIS FROM PIECE WHERE '+
                  'GP_NATUREPIECEG="'+NaturePieceG+'" AND '+
                  'GP_SOUCHE="'+SoucheG+'" AND '+
                  'GP_NUMERO='+InttoStr(Numero)+' AND '+
                  'GP_INDICEG='+InttoStr(IndiceG)+') AND '+
                  'BST_VIVANTE="X" ORDER BY BST_INDICESIT DESC';
    QQ := OpenSQL(Sql,true,1,'',true);
    if not QQ.Eof then
    begin
      Nature := QQ.fields[0].AsString;
      Souche := QQ.fields[1].AsString;
      Num    := QQ.fields[2].AsInteger;
      Indice := 0;
    end;
    Ferme(QQ);
  end else
  begin
    QQ := OpenSQL('SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_INDICEG FROM PIECE WHERE GP_NATUREPIECEG="FBP" AND '+
                  'GP_AFFAIREDEVIS=('+
                  'SELECT GP_AFFAIREDEVIS FROM PIECE WHERE '+
                  'GP_NATUREPIECEG="'+NaturePieceG+'" AND '+
                  'GP_SOUCHE="'+SoucheG+'" AND '+
                  'GP_NUMERO='+InttoStr(Numero)+' AND '+
                  'GP_INDICEG='+InttoStr(IndiceG)+') AND '+
                  'GP_VIVANTE="X"',true,1,'',true);
    if not QQ.Eof then
    begin
      Nature := QQ.fields[0].AsString;
      Souche := QQ.fields[1].AsString;
      Num    := QQ.fields[2].AsInteger;
      Indice := QQ.fields[3].AsInteger;
    end;
    Ferme(QQ);
  end;
  if Nature <> '' then
  begin
    TOBMF := TOB.create ('PIECE',TOBAFACTURER,-1);
    QQ := OpenSql ('SELECT PIECE.*,AFF_GENERAUTO,AFF_OKSIZERO,AFF_MANDATAIRE FROM PIECE '+
                   'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIREDEVIS '+
                   'WHERE '+
                   'GP_NATUREPIECEG="'+Nature+'" AND '+
                   'GP_SOUCHE="'+Souche+'" AND '+
                   'GP_NUMERO='+IntToStr(Num)+' AND '+
                   'GP_INDICEG='+IntToStr(indice),true,1,'',true);
    if not QQ.eof then
    begin
      TOBMF.SelectDB ('',QQ);
      TOBAfacturer.AddChampSupValeur ('GP_AFFAIRE',TOBMF.GetValue('GP_AFFAIRE'));
      TOBAfacturer.AddChampSupValeur ('GP_TIERS',TOBMF.GetValue('GP_TIERS'));
      TOBAfacturer.AddChampSupValeur ('LIBELLEAFFAIRE',GetChampsAffaire(TOBMF.GetValue('GP_AFFAIRE'),'AFF_LIBELLE'));
      TOBAfacturer.AddChampSupValeur ('LIBELLETIERS',RecupLibelleTiers(TOBMF.GetValue('GP_TIERS')));
      TOBAfacturer.AddChampSupValeur ('AFF_GENERAUTO',TOBMF.GetValue('AFF_GENERAUTO'));
      TOBAfacturer.AddChampSupValeur ('DATECLOTURE',DateCloture);
      TOBAfacturer.AddChampSupValeur ('AFF_MANDATAIRE',TOBMF.GetValue('AFF_MANDATAIRE'));
    end;
    ferme(QQ);
  end;
end;

function ExisteSituationProv (Avancement : Boolean; NaturePieceg,Souche: string; Numero,IndiceG : integer;
                                               var NatS,SoucheS : string; var NumS,indiceS : integer) : boolean;
var QQ : TQuery;
    SQl,RefDossier,SqlDoc : String;
begin
  NatS := ''; SoucheS := '';
  NumS := 0; IndiceS := 0;
  if Avancement then
  begin
    Result := false;
    // on recup le devis principal de factureation s'il existe bien sur un dossier de facturation
    SQl := 'SELECT M1.BMF_NATUREPIECEG,M1.BMF_SOUCHE,M1.BMF_NUMERO,M1.BMF_INDICEG FROM BTMEMOFACTURE M1 WHERE '+
           'M1.BMF_DEVISPRINC = '+
            '('+'SELECT M2.BMF_DEVISPRINC FROM BTMEMOFACTURE M2 WHERE '+
            'M2.BMF_NATUREPIECEG="'+NaturePieceg+'" AND '+
            'M2.BMF_SOUCHE="'+Souche+'" AND '+
            'M2.BMF_NUMERO='+InttOstr(Numero)+' AND '+
            'M2.BMF_INDICEG='+InttOstr(IndiceG)+
            ') ';

    QQ := OpenSql (SQl,true);
    if not QQ.eof then
    begin
      NatS := QQ.Fields[0].AsString;
      SoucheS := QQ.Fields[1].AsString;
      NumS := QQ.Fields[2].AsInteger;
      IndiceS := QQ.Fields[3].AsInteger;
    end;
    ferme (QQ);
    if (NatS = '') then Exit;
    SQlDoc := 'SELECT GP_AFFAIREDEVIS FROM PIECE WHERE '+
              'GP_NATUREPIECEG="'+NatS+'" AND '+
              'GP_SOUCHE="'+SoucheS+'" AND '+
              'GP_NUMERO='+IntToStr(NumS)+' AND '+
              'GP_INDICEG='+InttoStr(IndiceS);

    QQ := OpenSQL('SELECT 1 FROM BSITUATIONS WHERE BST_NATUREPIECE="FBP" AND '+
                  'BST_SSAFFAIRE=('+SQLDoc+') AND '+
                  'BST_VIVANTE="X"',true,1,'',true);
    Result := (not QQ.Eof);
    Ferme(QQ);
  end else
  begin
    NatS := NaturePieceG;
    SoucheS := Souche;
    NumS :=  Numero;
    IndiceS := IndiceG;
    //
    QQ := OpenSQL('SELECT 1 FROM PIECE P1 WHERE P1.GP_NATUREPIECEG="FBP" AND '+
                  'P1.GP_AFFAIREDEVIS=('+
                  'SELECT P2.GP_AFFAIREDEVIS FROM PIECE P2 WHERE P2.GP_NATUREPIECEG="'+NaturePieceG+'" AND '+
                  'P2.GP_SOUCHE="'+Souche+'" AND '+
                  'P2.GP_NUMERO='+InttoStr(Numero)+' AND '+
                  'P2.GP_INDICEG='+InttoStr(IndiceG)+') AND '+
                  'P1.GP_VIVANTE="X"',true,1,'',true);
    Result := (not QQ.Eof);
    Ferme(QQ);
  end;
end;

function TOF_BTSELFAC_MUL.AcompteNonRegle(Naturepieceg, Souche: string;numero, indiceg: integer): boolean;
var NonOk,pass : boolean;
begin
  result := false;
  If (ExisteSql ('SELECT 1 FROM PIECE WHERE GP_ATTACHEMENT<>"" AND '+
                      'GP_NATUREPIECEG="'+naturePieceg+'" AND '+
                      'GP_SOUCHE="'+Souche+'" AND '+
                      'GP_NUMERO='+IntToStr(Numero)+' AND '+
                      'GP_INDICEG='+IntToStr(Indiceg))) then
  begin
    result := not ExisteSql ('SELECT 1 FROM ACOMPTES WHERE GAC_JALECR="" AND '+
                        'GAC_NATUREPIECEG="'+naturePieceg+'" AND '+
                        'GAC_SOUCHE="'+Souche+'" AND '+
                        'GAC_NUMERO='+IntToStr(Numero)+' AND '+
                        'GAC_INDICEG='+IntToStr(Indiceg));
  end;
end;

Initialization
  registerclasses ( [ TOF_BTSELFAC_MUL ] ) ;
end.
