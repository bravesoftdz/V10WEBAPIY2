unit UTofPiece; 

interface
uses  M3FP, StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOB,LookUp,HDimension,Graphics,UtilGC,Ent1,
{$IFDEF EAGLCLIENT}
      MaineAgl,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,Fe_Main, MajTable,
{$ENDIF}
{$IFNDEF CCS3}
      UtilConfid,
{$ENDIF}
{$IFDEF AFFAIRE}
      AffaireUtil,
{$ENDIF}
{$IFDEF BTP}
      CalcOleGenericBTP ,
{$ENDIF}
      UTOF, AglInit, Agent,EntGC,Menus,TiersUtil,HTB97,FactUtil,ParamSoc,SaisUtil
      ,wCommuns 
      ;

Type

     TOF_ComplLigne = Class (TOF)
       private
         TobLigneAvantModif : tob;   
       procedure RechAffaire(Sender: TObject);
       Procedure AfficheDimensions(TOBArt : TOB);
       {$IFDEF MONCODE}
       Function  NbreDecimalSupp : Integer;
       procedure AffichageDecimalePrix;
       {$ENDIF}
       public
       procedure GL_DEPOTClick (Sender : TObject);
       procedure OnCancel ; override ;
       procedure OnNew ; override ;
       procedure OnDelete ; override ;
       procedure OnUpdate ; override ;
       procedure OnLoad ; override ;
       procedure OnArgument(stArgument : String ) ; override ;

     END ;

     TOF_ComplPiece = Class (TOF)
       private
       	 ModifAvanc : boolean;
         TOBAdresses : TOB;
         sTiersLivreOnEnter, sTiersFactureOnEnter : string;
         PutEcran: Boolean;
         NumeroContactEnter: string;
         ValuesContact : MyArrayValue;
         procedure SetAdressefacturationEnabled(Saisissable: boolean);
         procedure SetLastError (Num : integer; ou : string );

       public
       procedure OnArgument(stArgument : String ) ; override ;
       procedure OnCancel ; override ;
       procedure OnUpdate ; override ;
       procedure OnLoad ; override ;

       procedure GP_REFINTERNEOnExit(Sender: TObject); 
       procedure GP_REFEXTERNEOnExit(Sender: TObject); 

       procedure GP_TIERSLIVREOnEnter(Sender: TObject);   
       procedure GP_TIERSLIVREOnExit(Sender: TObject);    

       procedure GP_TIERSFACTUREOnEnter(Sender: TObject); 
       procedure GP_TIERSFACTUREOnExit(Sender: TObject);  

       procedure ADL_CODEPOSTALOnChange(Sender: TObject);  
       procedure ADL_VILLEOnChange(Sender: TObject);       
       procedure ADF_CODEPOSTALOnChange(Sender: TObject);  
       procedure ADF_VILLEOnChange(Sender: TObject);       

       Procedure ADL_PAYSOnExit(Sender: tObject); 
       Procedure ADF_PAYSOnExit(Sender: tObject); 

       procedure NUMEROCONTACTOnEnter(Sender: TObject); 
       procedure NUMEROCONTACTOnExit(Sender: TObject);  

       procedure CP_ADR_IDENTOnClick(Sender: TObject);

       procedure RechercheAdresseLivraison;           
       procedure RechercheAdresseFacturation;         
       procedure SetValeursContact(Prefixe: string);  

       procedure PutEcranTobPieceAdresseLivraison( TobAdresseLivraison : tob);
       procedure PutEcranTobPieceAdresseFacturation( TobAdresseFacturation : tob);  
       Function GetCodeAuxi(Prefixe: string) : string;                              

     END ;

     TOF_SaisCond = Class (TOF)
       private
       TOBCond : TOB;

       public
       procedure OnCancel ; override ;
       procedure OnNew ; override ;
       procedure OnDelete ; override ;
       procedure OnUpdate ; override ;
       procedure OnLoad ; override ;
     END ;

Const
  sTarifFournisseur       = '101';   //Fonctionnalité : tarif fournisseur
  sCommissionFournisseur  = '102';   //Fonctionnalité : commissionnement fournisseur
  sIndirectFournisseur    = '103';   //Fonctionnalité : indirect fournisseur

  sTarifClient            = '201';   //Fonctionnalité : tarif client
  sCommissionClient       = '202';   //Fonctionnalité : commissionnement client
  sIndirectClient         = '203';   //Fonctionnalité : indirect client

  sTarifSousTraitantAchat = '301';   //Fonctionnalité : tarif sous-traitant d'achat

  sTarifSousTraitantPhase = '401';   //Fonctionnalité : tarif sous-traitant de phase

const
	// libellés des messages
  TexteMessage: array[1..7] of string  = (
          {1}        'Vous devez renseigner un code tiers valide'
          {2}       ,'La raison sociale de l''adresse de livraison est obligatoire'
          {3}       ,'Le code postal de l''adresse de livraison est obligatoire'
          {4}       ,'La ville de l''adresse de livraison est obligatoire'
          {5}       , 'La raison sociale de l''adresse de facturation est obligatoire'
          {6}       ,'Le code postal de l''adresse de facturation est obligatoire'
          {7}       ,'La ville de l''adresse de facturation est obligatoire'
          );

procedure CalcQte (F : TForm; QteRef : String) ;
Procedure AdapteCond (F : TForm; QteArt : double) ;

Var ComplPieceCodeLivre, ComplPieceCodeFacture : string ;
    bAdrOk : boolean;

(*Const
	// libellés des messages
	TexteMessage: array[1..1] of string 	= (
          {1}         'Vous devez renseigner un code tiers valide.'
                 );*)


implementation
uses
  FactTiers,
  FactAdresse,
  BtpUtil,
  HrichOle
  ;
/////////////////////////////////////////////////////////////////////////////
procedure TOF_ComplLigne.OnCancel;
begin
  Inherited;
  if (TobLigneAvantModif<>nil) then TobLigneAvantModif.Free;  
end;

procedure TOF_ComplLigne.OnNew;
begin
    Inherited;
end;

procedure TOF_ComplLigne.OnDelete;
begin
    Inherited;
end;

procedure TOF_ComplLigne.OnLoad;
var
  TobArticle : TOB;
  sAchatOuVente : string;
begin
Inherited;
TobArticle := TOB.Create('ARTICLE',nil,-1);
TobArticle.SelectDB('"'+GetControlText('GL_ARTICLE')+'"',nil);
AfficheDimensions(TobArticle);
TobArticle.Free;

sAchatOuVente := GetInfoParPiece(LaTOB.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT');

IF LaTOB.GetValue('GL_ENCONTREMARQUE')='X' THEN
   BEGIN
   SetControlVisible ('GB_CONTREMARQUE', True);
   SetControlVisible ('CODEARTEREFERENCE', (LaTOB.GetValue('GL_TYPEREF')='CAT'));
   SetControlText('CODEARTEREFERENCE',Copy(LaTOB.GetValue('_CONTREMARTREF'),1,17));
   TCheckbox(GetControl('ARTNONREFERENCE')).checked:=(LaTOB.GetValue('GL_TYPEREF')='CAT');
   if GetInfoParPiece(LaTOB.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT')='ACH' THEN
      BEGIN
      SetControlCaption('TGL_FOURNISSEUR','Client');
      SetControlEnabled('BCLIENT',true);
      END ELSE SetControlEnabled('BFOUR',true);
   END ELSE
   BEGIN
   SetControlVisible ('GB_CONTREMARQUE', False);
   END;
{$IFDEF AFFAIRE}
if ctxAffaire in V_PGI.PGIContexte then
    BEGIN
    ChargeCleAffaire (Nil,THEDIT(GetControl('GL_AFFAIRE1')), THEDIT(GetControl('GL_AFFAIRE2')),
    THEDIT(GetControl('GL_AFFAIRE3')),THEDIT(GetControl('GL_AVENANT')),Nil, taCreat,GetControlText('GL_AFFAIRE'),False);
    END;
{$endif}
PopZoom97(TToolbarButton97(GetControl('BZOOM')),TPopupMenu(GetControl('POPMENU'))) ;
{$IFNDEF CCS3}
AppliquerConfidentialite(Ecran,GetInfoParPiece(LaTOB.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT'));
{$ENDIF}
{$IFDEF MONCODE}
AffichageDecimalePrix;
{$ENDIF}

  if       (sAchatOuVente='ACH') then
  begin
    if Assigned(GetControl('GL_TARIFSPECIAL')) then SetControlProperty('GL_TARIFSPECIAL', 'Plus', 'YTP_FONCTIONNALITE="'+sTarifFournisseur+'"');
    if Assigned(GetControl('GL_COMMSPECIAL' )) then SetControlProperty('GL_COMMSPECIAL' , 'Plus', 'YTP_FONCTIONNALITE="'+sCommissionFournisseur+'"');
  end
  else if  (sAchatOuVente='VEN') then
  begin
    if Assigned(GetControl('GL_TARIFSPECIAL')) then SetControlProperty('GL_TARIFSPECIAL', 'Plus', 'YTP_FONCTIONNALITE="'+sTarifClient+'"');
    if Assigned(GetControl('GL_COMMSPECIAL' )) then SetControlProperty('GL_COMMSPECIAL' , 'Plus', 'YTP_FONCTIONNALITE="'+sCommissionClient+'"');
  end
  else
  begin
    if Assigned(GetControl('GL_TARIFSPECIAL')) then SetControlProperty('GL_TARIFSPECIAL', 'Plus', 'YTP_FONCTIONNALITE="ZZZ"');
    if Assigned(GetControl('GL_COMMSPECIAL' )) then SetControlProperty('GL_COMMSPECIAL' , 'Plus', 'YTP_FONCTIONNALITE="ZZZ"');
  end;

if (VH_GC.GCIfDefCEGID) then
begin
  SetControlProperty ('GL_REFCOLIS', 'Enabled', False);
  SetControlProperty ('GL_RESSOURCE', 'Enabled', False);
  if GetControlText ('GL_REFCOLIS') <> '' then
    SetControlProperty ('GL_DATELIVRAISON', 'Enabled', False);
end;

end;

procedure TOF_ComplLigne.OnUpdate;
var
  NomChamp  : string;
  iCpt      : integer;
  lRechercheTarifaireAFaire : boolean;
  lRechercheCommissionAFaire : boolean;
  lCalculHTDeLaLigne        : boolean;      
begin
    Inherited;
// Contrôle de la ressource saisie à la ligne
  if (GetControlText('GL_RESSOURCE') <> '' ) then
     if Not(LookupValueExist(THEdit(Getcontrol('GL_RESSOURCE')))) then
        BEGIN
        PGIBox('Vous devez renseigner une ressource valide',Ecran.Caption);
        SetFocusControl('GL_RESSOURCE') ;
        TForm(Ecran).ModalResult:=0;
        Exit ;
        END;
{$IFNDEF CCS3}
NomChamp:=VerifierChampsObligatoires(Ecran,GetInfoParPiece(LaTOB.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT'));
   if NomChamp<>'' then
      begin
      NomChamp:=ReadTokenSt(NomChamp);
      SetFocusControl(NomChamp) ;
      PGIBox('La saisie du champs suivant est obligatoire : '+champToLibelle(NomChamp),Ecran.Caption);
      TForm(Ecran).ModalResult:=0;
      Exit ;
      end;
{$ENDIF}

  {
  Identification des modifications apportées en saisie par comparaison TobAvant et TobAprès
  afin de mettre à jour les champs de la Tob RECALCULTARIF et GL_RECALCULER
  et ainsi relancer les calculs au retour dans Facture.pas
  }
  if (TobLigneAvantModif<>nil) and (LaTob<>nil) then
  begin
    lRechercheTarifaireAFaire := False;
    lCalculHTDeLaLigne := False;
    iCpt:=0;
    while (iCpt < TobLigneAvantModif.NbChamps -1)  and ((not lRechercheTarifaireAFaire) {or (not lRechercheCommissionAFaire)} ) do
    begin
      Inc(iCpt);
      // On doit relancer le calcul du H.T. de la ligne
      if  (     ( not lCalculHTDeLaLigne )
            and (    (TobLigneAvantModif.GetNomChamp(iCpt)='GL_VALEURFIXEDEV')
                  or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_VALEURREMDEV')
                  or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_PUHTDEV')
                  or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_REMISELIBRE1')
                  or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_REMISELIBRE3')
                  or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_REMISELIBRE2')
                )
          ) then
      begin
        lCalculHTDeLaLigne := (TobLigneAvantModif.GetValeur(iCpt) <> LaTob.GetValeur(iCpt));
      end;

      // On doit relancer la recherche du tarif et donc le recalcul du H.T. de la ligne
      if  (     ( not lRechercheTarifaireAFaire )
            and (    (TobLigneAvantModif.GetNomChamp(iCpt)='GL_TARIFTIERS')
                  or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_TARIFARTICLE')
                  or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_TARIFSPECIAL')
                  or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_DATELIVRAISON')
                  or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_BLOQUETARIF')
                )
          ) then
      begin
        lRechercheTarifaireAFaire := (TobLigneAvantModif.GetValeur(iCpt) <> LaTob.GetValeur(iCpt));
        if  ( not lCalculHTDeLaLigne ) then lCalculHTDeLaLigne := (TobLigneAvantModif.GetValeur(iCpt) <> LaTob.GetValeur(iCpt));
      end;

      // On doit relancer la recherche du commissionnement
      if  (     ( not lRechercheCommissionAFaire )
            and (    (TobLigneAvantModif.GetNomChamp(iCpt)='GL_TARIFTIERS')
                  or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_TARIFARTICLE')
                  or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_COMMSPECIAL')
                  or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_DATELIVRAISON')
                  or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_REPRESENTANT')
           //       or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_REPRESENTANT2')
           //       or (TobLigneAvantModif.GetNomChamp(iCpt)='GL_REPRESENTANT3')
                )
          ) then
      begin
        lRechercheCommissionAFaire := (TobLigneAvantModif.GetValeur(iCpt) <> LaTob.GetValeur(iCpt));
      end;

    end;
    if (lRechercheTarifaireAFaire)  then LaTob.PutValue('RECALCULTARIF','X') else LaTob.PutValue('RECALCULTARIF','-');
    if (lCalculHTDeLaLigne)         then LaTob.PutValue('GL_RECALCULER','X') else LaTob.PutValue('GL_RECALCULER','-');
    if (lRechercheCommissionAFaire) then LaTob.PutValue('RECALCULCOMM' ,'X') else LaTob.PutValue('RECALCULCOMM' ,'-');
  end;
  if (TobLigneAvantModif<>nil) then TobLigneAvantModif.Free;

  // La tob est transmise à la saisie de pièce
  if (LaTOB<>Nil) then TheTob := LaTOB ;
end;

procedure TOF_ComplLigne.OnArgument(stArgument : String );
var
  BRechAff : TToolBarButton97;
  Color    : Tcolor;
  lEnabled, lVisible : boolean;
begin
    Inherited;
if ctxAffaire in V_PGI.PGIContexte then
    BEGIN
    BRechAff := TToolBarButton97(GetControl('BRECHAFFAIRE'));
    if (BRechAff<>nil) then BRechAff.OnClick:=RechAffaire;
    END;

if Not(ctxGCAFF in V_PGI.PGIContexte) then SetControlVisible ('GB_RESSOURCE', False);
GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GL_LIBREART', 10, '');
if pos ('FORCEINFO',stArgument) >0 then
    TPageControl(GetControl('PAGE')).ActivePage:=TTabSheet(getcontrol('INFORMATION'));
if LaTOB<>Nil then
   begin
   SetControlProperty('GL_REPRESENTANT', 'DataType', 'GCCOMMERCIAL') ;
   SetControlProperty('GL_REPRESENTANT', 'Plus', GetWhereRepresentant(LaTOB, '', True)+' ORDER BY GCL_LIBELLE') ;
   end;

  {$IFDEF MODE}
    SetControlEnabled('GL_QTESTOCK', False) ;
    SetControlProperty('GL_QTESTOCK', 'Color', clBtnFace) ;
    {$IFDEF FOS5}
      SetControlVisible('GBTARIF', False) ;
      SetControlVisible('GB_ACHAT', False) ;
      SetControlVisible('GB_MARGE', False) ;
    {$ENDIF FOS5}
  {$ENDIF MODE}

if Assigned(GetControl('GL_TARIFSPECIAL')) then SetControlProperty('GL_TARIFSPECIAL', 'Plus', 'YTP_FONCTIONNALITE="'+sTarifClient+'"');
if Assigned(GetControl('GL_COMMSPECIAL' )) then SetControlProperty('GL_COMMSPECIAL' , 'Plus', 'YTP_FONCTIONNALITE="'+sCommissionClient+'"');   

{ Mémorisation des données d'entrée pour pouvoir comparer en sortir et mette à jour les flags pour relancer certains calculs }
TobLigneAvantModif  := Tob.Create('_TobLigneAvant_', nil, -1);
TobLigneAvantModif.Dupliquer(LaTob, True, True);               

  if not (ctxGescom in V_PGI.PGIContexte) then
  begin
    SetControlEnabled ('GL_DEPOT', False);
  end else
  begin
    THValComboBox(GetControl('GL_DEPOT')).OnClick := GL_DEPOTClick;
  end;

  lVisible := GetParamSoc('SO_PREFSYSTTARIF');
  lEnabled := lVisible and (pos('CONSULTATION',stArgument)=0);
  if lEnabled then Color := clWindow else Color := clBtnFace;
  SetControlProperty('TGL_VALEURFIXEDEV'  , 'Visible', lVisible); SetControlProperty('GL_VALEURFIXEDEV'   , 'Visible', lVisible); SetControlProperty('GL_VALEURFIXEDEV', 'Enabled', lEnabled); SetControlProperty('GL_VALEURFIXEDEV', 'Color'  , Color);
  SetControlProperty('TGL_VALEURREMDEV'   , 'Visible', lVisible); SetControlProperty('GL_VALEURREMDEV'    , 'Visible', lVisible); SetControlProperty('GL_VALEURREMDEV' , 'Enabled', lEnabled); SetControlProperty('GL_VALEURREMDEV' , 'Color'  , Color);
  SetControlProperty('TGL_REMISELIBRE1'   , 'Visible', lVisible); SetControlProperty('GL_REMISELIBRE1'    , 'Visible', lVisible); SetControlProperty('GL_REMISELIBRE1' , 'Enabled', lEnabled); SetControlProperty('GL_REMISELIBRE1' , 'Color'  , Color);
  SetControlProperty('TGL_REMISELIBRE2'   , 'Visible', lVisible); SetControlProperty('GL_REMISELIBRE2'    , 'Visible', lVisible); SetControlProperty('GL_REMISELIBRE2' , 'Enabled', lEnabled); SetControlProperty('GL_REMISELIBRE2' , 'Color'  , Color);
  SetControlProperty('TGL_REMISELIBRE3'   , 'Visible', lVisible); SetControlProperty('GL_REMISELIBRE3'    , 'Visible', lVisible); SetControlProperty('GL_REMISELIBRE3' , 'Enabled', lEnabled); SetControlProperty('GL_REMISELIBRE3' , 'Color'  , Color);
                                                                  SetControlProperty('GL_BLOQUETARIF'     , 'Visible', lVisible); SetControlProperty('GL_BLOQUETARIF'  , 'Enabled', lEnabled);
  SetControlProperty('TGL_TARIFSPECIAL'   , 'Visible', lVisible); SetControlProperty('GL_TARIFSPECIAL'    , 'Visible', lVisible); SetControlProperty('GL_TARIFSPECIAL' , 'Enabled', lEnabled); SetControlProperty('GL_TARIFSPECIAL' , 'Color'  , Color);
  SetControlProperty('TREMISETOTALSYSTEME', 'Visible', lVisible); SetControlProperty('REMISETOTALSYSTEME' , 'Visible', lVisible); SetControlProperty('GL_TARIFSPECIAL' , 'Enabled', lEnabled); SetControlProperty('GL_TARIFSPECIAL' , 'Color'  , Color);
  SetControlEnabled('TGL_TARIFCLIENT', lEnabled); //JT eQualité 10997
  SetControlEnabled('GL_TARIFTIERS', lEnabled); //JT eQualité 10997
  //
END;

{$IFDEF MONCODE}
Function TOF_ComplLigne.NbreDecimalSupp : Integer;
Var PPQ : Double;
    NbDec : integer;
begin
Result:=0; NbDec:=0;
if LaTOB.GetValue('GL_DECIMALPRIX')='X' then
  begin
  PPQ:=LaTOB.GetValue('GL_PRIXPOURQTE'); if PPQ<=0 then PPQ:=1 ;
  While PPQ>1 do
    begin
    PPQ:=arrondi((PPQ/10.0)-0.499,0);
    inc(NbDec);
    end;
  Result:=NbDec;
  end;
end;

procedure TOF_ComplLigne.AffichageDecimalePrix;
Var NbDec : integer ;
    St : string;
    Prix,PPQ : double ;
begin
if LaTOB.GetValue('GL_DECIMALPRIX')<>'X' then Exit;
SetControlVisible('GL_PRIXPOURQTE',False);
SetControlVisible('TGL_PRIXPOURQTE',False);
Nbdec:=NbreDecimalSupp;
if GetControlVisible('GBVENTEPIVOT') then
  begin
  SetControlVisible('XXGBVENTEPIVOT',True);
  SetControlVisible('GBVENTEPIVOT',False);
  SetControlEnabled('XXGBVENTEPIVOT',GetControlEnabled('GBVENTEPIVOT'));
  SetControlVisible('XXGL_PUHT',GetControlVisible('GL_PUHT'));
  SetControlEnabled('XXGL_PUHT',GetControlEnabled('GL_PUHT'));
  SetControlVisible('XXGL_PUHTNET',GetControlVisible('GL_PUHTNET'));
  SetControlEnabled('XXGL_PUHTNET',GetControlEnabled('GL_PUHTNET'));
  SetControlVisible('XXGL_MONTANTHT',GetControlVisible('GL_MONTANTHT'));
  SetControlEnabled('XXGL_MONTANTHT',GetControlEnabled('GL_MONTANTHT'));
  PPQ:=LaTOB.GetValue('GL_PRIXPOURQTE'); if PPQ<=0 then PPQ:=1 ;
  Prix:=LaTOB.GetValue('GL_PUHT')/PPQ ;
  St:=StrF00(Prix,V_PGI.OkDecV+Nbdec);
  SetControlText('XXGL_PUHT',St);
  Prix:=LaTOB.GetValue('GL_PUHTNET')/PPQ ;
  St:=StrF00(Prix,V_PGI.OkDecV+Nbdec);
  SetControlText('XXGL_PUHTNET',St);
  SetControlText('XXGL_MONTANTHT',GetControlText('GL_MONTANTHT'));
  end;
if GetControlVisible('GBPRIXDEVISE') then
  begin
  SetControlVisible('XXGBPRIXDEVISE',True);
  SetControlVisible('GBPRIXDEVISE',False);
  SetControlEnabled('XXGBPRIXDEVISE',GetControlEnabled('GBPRIXDEVISE'));
  SetControlVisible('XXGL_PUHTDEV',GetControlVisible('GL_PUHTDEV'));
  SetControlEnabled('XXGL_PUHTDEV',GetControlEnabled('GL_PUHTDEV'));
  SetControlVisible('XXGL_REMISELIGNE',GetControlVisible('GL_REMISELIGNE'));
  SetControlEnabled('XXGL_REMISELIGNE',GetControlEnabled('GL_REMISELIGNE'));
  SetControlVisible('XXGL_PUHTNETDEV',GetControlVisible('GL_PUHTNETDEV'));
  SetControlEnabled('XXGL_PUHTNETDEV',GetControlEnabled('GL_PUHTNETDEV'));
  SetControlVisible('XXGL_MONTANTHTDEV',GetControlVisible('GL_MONTANTHTDEV'));
  SetControlEnabled('XXGL_MONTANTHTDEV',GetControlEnabled('GL_MONTANTHTDEV'));
  PPQ:=LaTOB.GetValue('GL_PRIXPOURQTE'); if PPQ<=0 then PPQ:=1 ;
  Prix:=LaTOB.GetValue('GL_PUHTDEV')/PPQ ;
  St:=StrF00(Prix,2+Nbdec);
  SetControlText('XXGL_PUHTDEV',St);
  Prix:=LaTOB.GetValue('GL_PUHTNETDEV')/PPQ ;
  St:=StrF00(Prix,2+Nbdec);
  SetControlText('XXGL_PUHTNETDEV',St);
  SetControlText('XXGL_REMISELIGNE',GetControlText('GL_REMISELIGNE'));
  SetControlText('XXGL_MONTANTHTDEV',GetControlText('GL_MONTANTHTDEV'));
  end;
end;

{$ENDIF}

/// SPECIF AFFAIRE Sur l'entête de pièce
procedure TOF_ComplLigne.RechAffaire(Sender: TObject);
{$ifdef AFFAIRE}
var Aff0, Aff1, Aff2, Aff3, Avenant    :string;
    EditAff, EditAff1, EditAff2, EditAff3 : THEdit;
{$endif}
begin
{$IFDEF AFFAIRE}
Aff1:='';Aff2:='';Aff3:='';
EditAff  :=THEdit(GetControl ('GL_AFFAIRE'));
EditAff1 :=THEdit(GetControl ('GL_AFFAIRE1'));
editAff2 :=THEdit(GetControl ('GL_AFFAIRE2'));
EditAff3 :=THEdit(GetControl ('GL_AFFAIRE3'));
// Mul de sélection des affaires
{$IFDEF BTP}
BTPCodeAffaireDecoupe(EditAff.Text,Aff0,Aff1,Aff2,Aff3,Avenant,taCreat, false);
{$ELSE}
CodeAffaireDecoupe(EditAff.Text,Aff0,Aff1,Aff2,Aff3,Avenant,taCreat, false);
{$ENDIF}
EditAff1.Text:= Aff1;
EditAff2.Text:= Aff2;
EditAff3.Text:= Aff3;
{$endif}
end;

Procedure TOF_ComplLigne.AfficheDimensions(TOBArt : TOB);
var i_indDim : integer;
    GrilleDim,CodeDim,LibDim,StDim : string;
begin
StDim:='';
if TobArt.GetValue('GA_STATUTART') = 'DIM' then
  for i_indDim := 1 to MaxDimension do
    begin
    GrilleDim := TOBArt.GetValue ('GA_GRILLEDIM' + IntToStr (i_indDim));
    CodeDim := TOBArt.GetValue ('GA_CODEDIM' + IntToStr (i_indDim));
    LibDim := GCGetCodeDim (GrilleDim, CodeDim, i_indDim);
    if LibDim <> '' then
       if StDim='' then StDim:=StDim + LibDim
       else StDim := StDim + ' - ' + LibDim;
    end;
if stDim <> '' then stDim := 'Dimension(s) : ' + StDim;
SetControlText('TARTDIM',StDim);
end;

Procedure TOF_ComplLigne.GL_DEPOTClick (Sender : TObject);
var TSql : TQuery;
begin
  TSql := OpenSQL ('SELECT GQ_DPA, GQ_PMAP, GQ_PMRP, GQ_DPR FROM DISPO WHERE GQ_CLOTURE="-" AND ' +
                   'GQ_ARTICLE="' + LaTOB.GetValue ('GL_ARTICLE') + '" AND GQ_DEPOT="' +
                   GetControlText ('GL_DEPOT') + '"', False);
  if not TSql.Eof then
  begin
    SetControlText ('GL_DPA', StrF00(TSql.FindField ('GQ_DPA').AsFloat, V_PGI.OkDecV));
    SetControlText ('GL_DPR', StrF00(TSql.FindField ('GQ_DPR').AsFloat, V_PGI.OkDecV));
    SetControlText ('GL_PMAP', StrF00(TSql.FindField ('GQ_PMAP').AsFloat, V_PGI.OkDecV));
    SetControlText ('GL_PMRP', StrF00(TSql.FindField ('GQ_PMRP').AsFloat, V_PGI.OkDecV));
  end;
Ferme (TSql);
end;

////////////////////////////////////////////////////////////////////////////////
///          E N T E T E    d e    P I E C E         ///////////////////////////
////////////////////////////////////////////////////////////////////////////////

Procedure TOF_ComplPiece.ADL_PAYSOnExit(Sender: tObject);
Begin
  SetControlProperty('ADL_REGION','PLUS','RG_PAYS="'+GetControlText('ADL_PAYS')+'"');
End;

Procedure TOF_ComplPiece.ADF_PAYSOnExit(Sender: tObject);
Begin
  SetControlProperty('ADF_REGION','PLUS','RG_PAYS="'+GetControlText('ADF_PAYS')+'"');
End;

procedure TOF_ComplPiece.GP_REFEXTERNEOnExit(Sender: TObject);
begin
   if not ReferenceIntExtIsOk(crExterne, LaTob, GetControlText('GP_REFEXTERNE')) then
      ThEdit(GetControl('GP_REFEXTERNE')).SetFocus;
end;

procedure TOF_ComplPiece.GP_REFINTERNEOnExit(Sender: TObject);
begin
   if not ReferenceIntExtIsOk(crInterne, LaTob, GetControlText('GP_REFINTERNE')) then
      ThEdit(GetControl('GP_REFINTERNE')).SetFocus;
end;

{ Initialisation du cope pays en fonction du code postal et/ou de la ville }
procedure TOF_ComplPiece.ADL_CODEPOSTALOnChange(Sender: TObject);
var
  cPays : string;
begin
  cPays := GetFieldFromCodePostal( 'O_PAYS', GetControlText('ADL_CODEPOSTAL'), GetControlText('ADL_VILLE'));
  if cPays <>'' then SetControlText('ADL_PAYS', cPays);
end;

procedure TOF_ComplPiece.ADL_VILLEOnChange(Sender: TObject);
var
  cPays : string;
begin
  cPays := GetFieldFromCodePostal( 'O_PAYS', GetControlText('ADL_CODEPOSTAL'), GetControlText('ADL_VILLE'));
  if cPays <>'' then SetControlText('ADL_PAYS', cPays);
end;

procedure TOF_ComplPiece.ADF_CODEPOSTALOnChange(Sender: TObject);
var
  cPays : string;
begin
  cPays := GetFieldFromCodePostal( 'O_PAYS', GetControlText('ADF_CODEPOSTAL'), GetControlText('ADF_VILLE'));
  if cPays <>'' then SetControlText('ADF_PAYS', cPays);
end;

procedure TOF_ComplPiece.ADF_VILLEOnChange(Sender: TObject);
var
  cPays : string;
begin
  cPays := GetFieldFromCodePostal( 'O_PAYS', GetControlText('ADF_CODEPOSTAL'), GetControlText('ADF_VILLE'));
  if cPays <>'' then SetControlText('ADF_PAYS', cPays);
end;

{ Adresse livraison }
procedure TOF_ComplPiece.GP_TIERSLIVREOnEnter(Sender: TObject);
begin
   sTiersLivreOnEnter := GetControlText('GP_TIERSLIVRE');
end;

procedure TOF_ComplPiece.GP_TIERSLIVREOnExit(Sender: TObject);
begin
  if (sTiersLivreOnEnter<>GetControlText('GP_TIERSLIVRE')) then
  begin
    //SetControlText('GP_NADRESSELIV',IntToStr(GetNumAdresseFromTiers('', GetControlText('GP_TIERSLIVRE'), taLivr)));
    SetControlText('GP_NADRESSELIV', GetNumAdresseFromTiers('', GetControlText('GP_TIERSLIVRE'), taLivr));
    SetControlProperty('ADL_NUMEROCONTACT', 'Plus', GetCodeAuxi('ADL'));
    RechercheAdresseLivraison;
  end;
end;

(*
procedure TOF_ComplPiece.GP_NADRESSELIVOnChange(Sender: TObject);
begin
  if not PutEcran then
    RechercheAdresseLivraison;
end;
*)

{ Adresse Facturation }
procedure TOF_ComplPiece.GP_TIERSFACTUREOnEnter(Sender: TObject);
begin
   sTiersFactureOnEnter := GetControlText('GP_TIERSFACTURE');
end;

procedure TOF_ComplPiece.GP_TIERSFACTUREOnExit(Sender: TObject);
begin
  if (sTiersFactureOnEnter<>GetControlText('GP_TIERSFACTURE')) then
  begin
    //SetControlText('GP_NADRESSEFAC',IntToStr(GetNumAdresseFromTiers('', GetControlText('GP_TIERSFACTURE'), taFact)));
    SetControlText('GP_NADRESSEFAC', GetNumAdresseFromTiers('', GetControlText('GP_TIERSFACTURE'), taFact));
    SetControlProperty('ADF_NUMEROCONTACT', 'Plus', GetCodeAuxi('FAC'));
    RechercheAdresseFacturation;
  end;
end;

(*
procedure TOF_ComplPiece.GP_NADRESSEFACOnChange(Sender: TObject);
begin
  if not PutEcran then
    RechercheAdresseFacturation;
end;
*)

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 30/01/2003
Modifié le ... :   /  /
Description .. : Positionne le filtre sur les adresses de livraison en fonction
Suite ........ : du tiers livré
Mots clefs ... :
*****************************************************************}
(*
procedure TOF_ComplPiece.SetFiltreAdresseLivraison(sTiersLivre : string);
begin
   SetControlProperty('GP_NADRESSELIV', 'Plus',       'ADR_TYPEADRESSE="TIE"'
                                               + ' AND ADR_REFCODE="' + sTiersLivre + '"'
                                               + ' AND ADR_LIVR="X"');
end;
*)

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 30/01/2003
Modifié le ... :   /  /
Description .. : Recherche d'une adresse de livraison
Suite ........ : Alimentation de la tob Adresses et affichage de celle-ci
Mots clefs ... :
*****************************************************************}
procedure TOF_ComplPiece.RechercheAdresseLivraison;
var
   TobAdresse : tob;
begin
   //TobAdresse  est une Tob liée à la table ADRESSES
   //TobAdresses est une Tob liée à la table PIECEADRESSE (ou à la table ADRESSES : ancien fonctionnement)
   TobAdresse := Tob.create('_ADRESSES_',nil,-1);
   Try
      GetTobAdresseFromAdresses(TobAdresse, 'TIE', GetControlText('GP_TIERSLIVRE'), StrToInt(GetControlText('GP_NADRESSELIV')), 'Livraison');
      //Si on trouve une adresse
      if (TobAdresse.Detail.count<>0) then
      begin
         GetTobPieceAdresseFromTobAdresses( TobAdresses.Detail[0], TobAdresse.Detail[0]);
      end
      else // sinon on reprend l'adresse du tiers
      begin
         GetAdrFromCode(TOBAdresses.Detail[0], GetControlText('GP_TIERSLIVRE')) ;
      end;
      PutEcranTobPieceAdresseLivraison( TobAdresses.Detail[0] );

      // Bug 10538 //
      if TCheckBox(GetControl('CP_ADR_IDENT')).Checked then
      begin
        PutEcranTobPieceAdresseFacturation( TobAdresses.Detail[0] );
        SetControlText('GP_TIERSFACTURE', GetControlText('GP_TIERSLIVRE'));
      end;
   Finally
      TobAdresse.Free;
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 30/01/2003
Modifié le ... :   /  /
Description .. : Put Ecran de la Tob Adresse pièce (table ADRESSES ou
Suite ........ : table PIECEADRESSE ) pour l'adresse de livraison
Mots clefs ... :
*****************************************************************}
procedure TOF_ComplPiece.PutEcranTobPieceAdresseLivraison( TobAdresseLivraison : tob);
var
  sIncoterm : string;
begin
  PutEcran := True;
   if GetParamSoc('SO_GCPIECEADRESSE') then
   begin
      //Adresse de Livraison
      if TobAdresseLivraison.GetValue('GPA_LIBELLE')<>'' then SetControlText('ADL_LIBTIERS',TobAdresseLivraison.GetValue('GPA_LIBELLE'));
      SetControlText('ADL_LIBTIERS2' ,TobAdresseLivraison.GetValue('GPA_LIBELLE2')    );
      SetControlText('ADL_ADRESSE1'  ,TobAdresseLivraison.GetValue('GPA_ADRESSE1')    );
      SetControlText('ADL_ADRESSE2'  ,TobAdresseLivraison.GetValue('GPA_ADRESSE2')    );
      SetControlText('ADL_ADRESSE3'  ,TobAdresseLivraison.GetValue('GPA_ADRESSE3')    );
      SetControlText('ADL_CODEPOSTAL',TobAdresseLivraison.GetValue('GPA_CODEPOSTAL')  );
      SetControlText('ADL_VILLE'     ,TobAdresseLivraison.GetValue('GPA_VILLE')       );
      SetControlText('ADL_PAYS'      ,TobAdresseLivraison.GetValue('GPA_PAYS')        );
      SetControlProperty('ADL_REGION','PLUS','RG_PAYS="'+GetControlText('ADL_PAYS')+'"'); 
      SetControlText('ADL_INCOTERM'  ,TobAdresseLivraison.GetValue('GPA_INCOTERM' )   );  
      sIncoterm := RechDom('GCINCOTERM', TobAdresseLivraison.GetValue('GPA_INCOTERM' ) , False);
      If sIncoterm = 'Error' then sIncoterm := '';
      SetControlCaption('TTADL_INCOTERM', sIncoterm);
      SetControlText('ADL_MODEEXP'   ,TobAdresseLivraison.GetValue('GPA_MODEEXP'  )   );  
      SetControlText('ADL_LIEUDISPO' ,TobAdresseLivraison.GetValue('GPA_LIEUDISPO')   );  
      SetControlText('ADL_NIF'       ,TobAdresseLivraison.GetValue('GPA_NIF')         );  
      SetControlText('ADL_EAN'       ,TobAdresseLivraison.GetValue('GPA_EAN')         );  
      SetControlText('ADL_REGION'    ,TobAdresseLivraison.GetValue('GPA_REGION')      );  
      SetControlText('ADL_NUMEROCONTACT',TobAdresseLivraison.GetValue('GPA_NUMEROCONTACT')  ); 
      ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], GetCodeAuxi('ADL'), TobAdresseLivraison.GetValue('GPA_NUMEROCONTACT')); 
      SetValeursContact('ADL');
   end
   else
   begin
      //Adresse de Livraison
      if TobAdresseLivraison.GetValue('ADR_LIBELLE')<>'' then SetControlText('ADL_LIBTIERS',TobAdresseLivraison.GetValue('ADR_LIBELLE'));
      SetControlText('ADL_LIBTIERS2' ,TobAdresseLivraison.GetValue('ADR_LIBELLE2')    );
      SetControlText('ADL_ADRESSE1'  ,TobAdresseLivraison.GetValue('ADR_ADRESSE1')    );
      SetControlText('ADL_ADRESSE2'  ,TobAdresseLivraison.GetValue('ADR_ADRESSE2')    );
      SetControlText('ADL_ADRESSE3'  ,TobAdresseLivraison.GetValue('ADR_ADRESSE3')    );
      SetControlText('ADL_CODEPOSTAL',TobAdresseLivraison.GetValue('ADR_CODEPOSTAL')  );
      SetControlText('ADL_VILLE'     ,TobAdresseLivraison.GetValue('ADR_VILLE')       );
      SetControlText('ADL_PAYS'      ,TobAdresseLivraison.GetValue('ADR_PAYS')        );
      SetControlProperty('ADL_REGION','PLUS','RG_PAYS="'+GetControlText('ADL_PAYS')+'"');
      SetControlText('ADL_INCOTERM'  ,TobAdresseLivraison.GetValue('ADR_INCOTERM' )   ); 
      sIncoterm := RechDom('GCINCOTERM', TobAdresseLivraison.GetValue('ADR_INCOTERM' ) , False);
      If sIncoterm = 'Error' then sIncoterm := '';
      SetControlCaption('TTADL_INCOTERM', sIncoterm);
      SetControlText('ADL_MODEEXP'   ,TobAdresseLivraison.GetValue('ADR_MODEEXP'  )   ); 
      SetControlText('ADL_LIEUDISPO' ,TobAdresseLivraison.GetValue('ADR_LIEUDISPO')   ); 
      SetControlText('ADL_NIF'       ,TobAdresseLivraison.GetValue('ADR_NIF')         ); 
      SetControlText('ADL_EAN'       ,TobAdresseLivraison.GetValue('ADR_EAN')         ); 
      SetControlText('ADL_REGION'    ,TobAdresseLivraison.GetValue('ADR_REGION')      ); { GPAO1 REGION }
      SetControlText('ADL_NUMEROCONTACT'  ,TobAdresseLivraison.GetValue('ADR_NUMEROCONTACT') ); 
      ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], GetCodeAuxi('ADL'), TobAdresseLivraison.GetValue('ADR_NUMEROCONTACT'));
      SetValeursContact('ADL');
   end;
  PutEcran := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 30/01/2003
Modifié le ... :   /  /
Description .. : Positionne le filtre sur les adresses de facturation en
Suite ........ : fonction du tiers facturé
Mots clefs ... :
*****************************************************************}
(*
procedure TOF_ComplPiece.SetFiltreAdresseFacturation(sTiersFacture : string);
begin
   SetControlProperty('GP_NADRESSEFAC', 'Plus',       'ADR_TYPEADRESSE="TIE"'
                                               + ' AND ADR_REFCODE="' + sTiersFacture + '"'
                                               + ' AND ADR_FACT="X"');
end;
*)

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 30/01/2003
Modifié le ... :   /  /
Description .. : Recherche d'une adresse de facturation
Suite ........ : Alimentation de la tob Adresses et affichage de celle-ci
Mots clefs ... :
*****************************************************************}
procedure TOF_ComplPiece.RechercheAdresseFacturation;
var
   TobAdresse : tob;
begin
   //TobAdresse  est une Tob liée à la table ADRESSES
   //TobAdresses est une Tob liée à la table PIECEADRESSE (ou à la table ADRESSES : ancien fonctionnement)
   TobAdresse := Tob.create('_ADRESSES_',nil,-1);
   Try
      GetTobAdresseFromAdresses(TobAdresse, 'TIE', GetControlText('GP_TIERSFACTURE'), StrToInt(GetControlText('GP_NADRESSEFAC')), 'Facturation');
      //Si on trouve une adresse
      if (TobAdresse.Detail.count<>0) then
      begin
         GetTobPieceAdresseFromTobAdresses( TobAdresses.Detail[1], TobAdresse.Detail[0]);
      end
      else // sinon on reprend l'adresse du tiers
      begin
         GetAdrFromCode(TOBAdresses.Detail[1], GetControlText('GP_TIERSFACTURE')) ;
      end;
      PutEcranTobPieceAdresseFacturation( TobAdresses.Detail[1] );
   Finally
      TobAdresse.Free;
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 30/01/2003
Modifié le ... :   /  /
Description .. : Put Ecran de la Tob Adresse pièce (table ADRESSES ou
Suite ........ : table PIECEADRESSE ) pour l'adresse de facturation
Mots clefs ... :
*****************************************************************}
procedure TOF_ComplPiece.PutEcranTobPieceAdresseFacturation( TobAdresseFacturation : tob);
Var
  sIncoterm : string;
begin
   if GetParamSoc('SO_GCPIECEADRESSE') then
   begin
      //Adresse de Facturation
      if TobAdresseFacturation.GetValue('GPA_LIBELLE')<>'' then SetControlText('ADF_LIBTIERS',TobAdresseFacturation.GetValue('GPA_LIBELLE'));
      SetControlText('ADF_LIBTIERS2'    ,TobAdresseFacturation.GetValue('GPA_LIBELLE2')    );
      SetControlText('ADF_ADRESSE1'     ,TobAdresseFacturation.GetValue('GPA_ADRESSE1')    );
      SetControlText('ADF_ADRESSE2'     ,TobAdresseFacturation.GetValue('GPA_ADRESSE2')    );
      SetControlText('ADF_ADRESSE3'     ,TobAdresseFacturation.GetValue('GPA_ADRESSE3')    );
      SetControlText('ADF_CODEPOSTAL'   ,TobAdresseFacturation.GetValue('GPA_CODEPOSTAL')  );
      SetControlText('ADF_VILLE'        ,TobAdresseFacturation.GetValue('GPA_VILLE')       );
      SetControlText('ADF_PAYS'         ,TobAdresseFacturation.GetValue('GPA_PAYS')        );
      SetControlProperty('ADF_REGION','PLUS','RG_PAYS="'+GetControlText('ADF_PAYS')+'"');     
      SetControlText('ADF_INCOTERM'     ,TobAdresseFacturation.GetValue('GPA_INCOTERM' )   ); 
      sIncoterm := RechDom('GCINCOTERM', TobAdresseFacturation.GetValue('GPA_INCOTERM' ) , False);
      If sIncoterm = 'Error' then sIncoterm := '';
      SetControlCaption('TTADF_INCOTERM', sIncoterm);
      SetControlText('ADF_MODEEXP'      ,TobAdresseFacturation.GetValue('GPA_MODEEXP'  )   ); 
      SetControlText('ADF_LIEUDISPO'    ,TobAdresseFacturation.GetValue('GPA_LIEUDISPO')   ); 
      SetControlText('ADF_NIF'          ,TobAdresseFacturation.GetValue('GPA_NIF')         ); 
      SetControlText('ADF_EAN'          ,TobAdresseFacturation.GetValue('GPA_EAN')         ); 
      SetControlText('ADF_REGION'       ,TobAdresseFacturation.GetValue('GPA_REGION')      ); 
      SetControlText('ADF_NUMEROCONTACT',TobAdresseFacturation.GetValue('GPA_NUMEROCONTACT') ); 

      ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], GetCodeAuxi('FAC'), TobAdresseFacturation.GetValue('GPA_NUMEROCONTACT')); 
      SetValeursContact('ADF');
   end
   else
   begin
      //Adresse de Facturation
      if TobAdresseFacturation.GetValue('ADR_LIBELLE')<>'' then SetControlText('ADF_LIBTIERS',TobAdresseFacturation.GetValue('ADR_LIBELLE'));
      SetControlText('ADF_LIBTIERS2'    ,TobAdresseFacturation.GetValue('ADR_LIBELLE2')    );
      SetControlText('ADF_ADRESSE1'     ,TobAdresseFacturation.GetValue('ADR_ADRESSE1')    );
      SetControlText('ADF_ADRESSE2'     ,TobAdresseFacturation.GetValue('ADR_ADRESSE2')    );
      SetControlText('ADF_ADRESSE3'     ,TobAdresseFacturation.GetValue('ADR_ADRESSE3')    );
      SetControlText('ADF_CODEPOSTAL'   ,TobAdresseFacturation.GetValue('ADR_CODEPOSTAL')  );
      SetControlText('ADF_VILLE'        ,TobAdresseFacturation.GetValue('ADR_VILLE')       );
      SetControlText('ADF_PAYS'         ,TobAdresseFacturation.GetValue('ADR_PAYS')        );
      SetControlProperty('ADF_REGION','PLUS','RG_PAYS="'+GetControlText('ADF_PAYS')+'"');     
      SetControlText('ADF_INCOTERM'     ,TobAdresseFacturation.GetValue('ADR_INCOTERM' )   ); 
      sIncoterm := RechDom('GCINCOTERM', TobAdresseFacturation.GetValue('ADR_INCOTERM' ) , False);
      If sIncoterm = 'Error' then sIncoterm := '';
      SetControlCaption('TTADF_INCOTERM', sIncoterm);
      SetControlText('ADF_MODEEXP'      ,TobAdresseFacturation.GetValue('ADR_MODEEXP'  )   ); 
      SetControlText('ADF_LIEUDISPO'    ,TobAdresseFacturation.GetValue('ADR_LIEUDISPO')   ); 
      SetControlText('ADF_NIF'          ,TobAdresseFacturation.GetValue('ADR_NIF')         ); 
      SetControlText('ADF_EAN'          ,TobAdresseFacturation.GetValue('ADR_EAN')         ); 
      SetControlText('ADF_EAN'          ,TobAdresseFacturation.GetValue('ADR_EAN')         ); 
      SetControlText('ADF_REGION'       ,TobAdresseFacturation.GetValue('ADR_REGION')      ); 
      SetControlText('ADF_NUMEROCONTACT',TobAdresseFacturation.GetValue('ADR_NUMEROCONTACT') ); 
      ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], GetCodeAuxi('FAC'), TobAdresseFacturation.GetValue('ADR_NUMEROCONTACT')); 
      SetValeursContact('ADF');
   end;
end;

procedure TOF_ComplPiece.OnArgument(stArgument: String);
begin
   inherited;
	 AppliqueFontDefaut (THRichEditOle(GetControl('GP_BLOCNOTE')));
   if Assigned(GetControl('GP_REFINTERNE') ) then ThEdit(GetControl('GP_REFINTERNE')).OnExit := GP_REFINTERNEOnExit;
   if Assigned(GetControl('GP_REFEXTERNE') ) then ThEdit(GetControl('GP_REFEXTERNE')).OnExit := GP_REFEXTERNEOnExit;

   if Assigned(GetControl('GP_TIERSLIVRE')  ) then ThEdit(GetControl('GP_TIERSLIVRE')  ).OnEnter   := GP_TIERSLIVREOnEnter;
   if Assigned(GetControl('GP_TIERSLIVRE')  ) then ThEdit(GetControl('GP_TIERSLIVRE')  ).OnExit    := GP_TIERSLIVREOnExit;
   if Assigned(GetControl('GP_TIERSFACTURE')) then ThEdit(GetControl('GP_TIERSFACTURE')).OnEnter   := GP_TIERSFACTUREOnEnter;
   if Assigned(GetControl('GP_TIERSFACTURE')) then ThEdit(GetControl('GP_TIERSFACTURE')).OnExit    := GP_TIERSFACTUREOnExit;
//   if Assigned(GetControl('GP_NADRESSELIV') ) then ThEdit(GetControl('GP_NADRESSELIV') ).OnChange  := GP_NADRESSELIVOnChange;
//   if Assigned(GetControl('GP_NADRESSEFAC') ) then ThEdit(GetControl('GP_NADRESSEFAC') ).OnChange  := GP_NADRESSEFACOnChange;

   if Assigned(GetControl('ADL_CODEPOSTAL') ) then ThEdit(GetControl('ADL_CODEPOSTAL') ).OnChange  := ADL_CODEPOSTALOnChange;
   if Assigned(GetControl('ADL_VILLE')      ) then ThEdit(GetControl('ADL_VILLE')      ).OnChange  := ADL_VILLEOnChange;
   if Assigned(GetControl('ADF_CODEPOSTAL') ) then ThEdit(GetControl('ADF_CODEPOSTAL') ).OnChange  := ADF_CODEPOSTALOnChange;
   if Assigned(GetControl('ADF_VILLE')      ) then ThEdit(GetControl('ADF_VILLE')      ).OnChange  := ADF_VILLEOnChange;

   if Assigned(GetControl('ADL_PAYS')  ) then ThValComboBox(GetControl('ADL_PAYS')  ).OnExit   := ADL_PAYSOnExit;
   if Assigned(GetControl('ADF_PAYS')  ) then ThValComboBox(GetControl('ADF_PAYS')  ).OnExit   := ADF_PAYSOnExit;

   if Assigned(GetControl('ADL_NUMEROCONTACT')) then
   begin
     ThEdit(GetControl('ADL_NUMEROCONTACT')).OnEnter := NUMEROCONTACTOnEnter;
     ThEdit(GetControl('ADL_NUMEROCONTACT')).OnExit := NUMEROCONTACTOnExit;
   end;
   if Assigned(GetControl('ADF_NUMEROCONTACT')) then
   begin
     ThEdit(GetControl('ADF_NUMEROCONTACT')).OnEnter := NUMEROCONTACTOnEnter;
     ThEdit(GetControl('ADF_NUMEROCONTACT')).OnExit := NUMEROCONTACTOnExit;
   end;
  if Assigned(GetControl('CP_ADR_IDENT')) then
     ThEdit(GetControl('CP_ADR_IDENT')).OnClick := CP_ADR_IDENTOnClick;
{$IFDEF CCS3}
  SetControlVisible ('GBPIECE', False);
{$ENDIF}

end;

/////////////////////////////////////////////////////////////////////////////
procedure TOF_ComplPiece.OnCancel;
begin
    Inherited;
    TOBAdresses.Free;
end;

procedure TOF_ComplPiece.OnLoad;
Var F : TForm ;
    i_ind : integer;
    Titre,NatPiece,NatTiers,St,Where,WhereCli : string;
begin
F:=TForm(Ecran);
NatPiece:=LaTOB.GetValue('GP_NATUREPIECEG') ;
bAdrOk := false;
TobAdresses:=TOB(LaTob.data);
if ((LaTOB.GetValue('GP_TOTALTTCDEV')=0) and (LaTOB.GetValue('GP_TOTALQTEFACT')=0)) then
    BEGIN
    SetControlEnabled('GP_REGIMETAXE',True);
    SetControlProperty('GP_REGIMETAXE','COLOR',clWindow);
    END;

   if       (LaTob.GetValue('GP_VENTEACHAT')='ACH') then
   begin
     if Assigned(GetControl('GP_TARIFSPECIAL')) then SetControlProperty('GP_TARIFSPECIAL', 'Plus', 'YTP_FONCTIONNALITE="'+sTarifFournisseur+'"');
     if Assigned(GetControl('GP_COMMSPECIAL' )) then SetControlProperty('GP_COMMSPECIAL' , 'Plus', 'YTP_FONCTIONNALITE="'+sCommissionFournisseur+'"');
   end
   else if  (LaTob.GetValue('GP_VENTEACHAT')='VEN') then
   begin
     if Assigned(GetControl('GP_TARIFSPECIAL')) then SetControlProperty('GP_TARIFSPECIAL', 'Plus', 'YTP_FONCTIONNALITE="'+sTarifClient+'"');
     if Assigned(GetControl('GP_COMMSPECIAL' )) then SetControlProperty('GP_COMMSPECIAL' , 'Plus', 'YTP_FONCTIONNALITE="'+sCommissionClient+'"');
   end
   else
   begin
     if Assigned(GetControl('GP_TARIFSPECIAL')) then SetControlProperty('GP_TARIFSPECIAL', 'Plus', 'YTP_FONCTIONNALITE="ZZZ"');
     if Assigned(GetControl('GP_COMMSPECIAL' )) then SetControlProperty('GP_COMMSPECIAL' , 'Plus', 'YTP_FONCTIONNALITE="ZZZ"');
   end;

if LaTob.GetValue('GP_VENTEACHAT') = 'VEN' then
   begin
   SetControlText('TGP_TIERS','Client');
   SetControlProperty('GP_TIERSLIVRE','DATATYPE','GCTIERSCLI');
   SetControlProperty('GP_TIERSFACTURE','DATATYPE','GCTIERSCLI');
   for i_ind:=1 to 9 do SetControlProperty('GP_LIBRETIERS'+IntToStr(i_ind),'DATATYPE','GCLIBRETIERS'+IntToStr(i_ind));
   SetControlProperty('GP_LIBRETIERSA','DATATYPE','GCLIBRETIERSA');
   GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GP_LIBRETIERS', 10, '');
   end
   else if LaTob.GetValue('GP_VENTEACHAT') = 'ACH' then
   begin
   SetControlText('TGP_TIERS','Fournisseur');
// BBI, correction fiche 10374
   SetControlProperty('GP_TIERSLIVRE','DATATYPE','GCTIERSFOURN'); //Modif du 12/06/01
   SetControlProperty('GP_TIERSFACTURE','DATATYPE','GCTIERSFOURN');
// BBI, fin correction fiche 10374
   for i_ind:=1 to 3 do SetControlProperty('GP_LIBRETIERS'+IntToStr(i_ind),'DATATYPE','GCLIBREFOU'+IntToStr(i_ind));
   for i_ind:=4 to 9 do
    begin
    SetControlVisible('TGP_LIBRETIERS'+IntToStr(i_ind),False);
    SetControlVisible('GP_LIBRETIERS'+IntToStr(i_ind),False);
    end ;
   SetControlVisible('TGP_LIBRETIERSA',False);
   SetControlVisible('GP_LIBRETIERSA',False);
   for i_ind:=1 to 3 do
    BEGIN
    GCTitreZoneLibre('YTC_TABLELIBREFOU'+IntToStr(i_ind),Titre);
    if Titre = '' then
        begin
        SetControlVisible('TGP_LIBRETIERS'+IntToStr(i_ind),False);
        SetControlVisible('GP_LIBRETIERS'+IntToStr(i_ind),False);
        end else SetControlCaption('TGP_LIBRETIERS'+IntToStr(i_ind), Titre) ;
    END;
   SetControlProperty('GBCLIENT', 'CAPTION','Fournisseur') ;
   end;

ComplPieceCodeLivre:=LaTOB.GetValue('GP_TIERSLIVRE');
ComplPieceCodeFacture:=LaTOB.GetValue('GP_TIERSFACTURE');
LaTob.PutEcran(F);
NatTiers:=GetInfoParPiece(NatPiece,'GPP_NATURETIERS');
Where:=''; WhereCli:=''; St:=NatTiers;
While St <> '' do
      begin
      if WhereCli='' then WhereCli:=' T_NATUREAUXI="'+ReadTokenSt(St)+'"'
                     else WhereCli:=WhereCli+' OR T_NATUREAUXI="'+ReadTokenSt(St)+'"';
      end;
if WhereCli<>'' then Where:=' AND ('+WhereCli+')';
if not ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+
     LaTob.GetValue('GP_TIERSLIVRE') + '"' + Where) then
    begin
    SetControlText('GP_TIERSLIVRE','');
    SetControlText('GP_TIERSFACTURE','');
    exit;
    end;

if LaTOB.GetValue('GP_TIERS')=''  then SetControlText('LIBTIERS','')
else SetControlText('LIBTIERS',RechDom('GCTIERS',LaTOB.GetValue('GP_TIERS'),FALSE));

if LaTOB.GetValue('GP_TIERSLIVRE')=''  then SetControlText('ADL_LIBTIERS','')
else
begin
  SetControlText('ADL_LIBTIERS',RechDom('GCTIERS',LaTOB.GetValue('GP_TIERSLIVRE'),FALSE));
  SetControlProperty('ADL_NUMEROCONTACT', 'Plus', GetCodeAuxi('ADL'));

  ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], GetCodeAuxi('ADL'), ValeurI(GetControlText('ADL_NUMEROCONTACT')));
  SetValeursContact('ADL');
end;

if LaTOB.GetValue('GP_TIERSFACTURE')=''  then SetControlText('ADF_LIBTIERS','')
else
begin
  SetControlText('ADF_LIBTIERS',RechDom('GCTIERS',LaTOB.GetValue('GP_TIERSFACTURE'),FALSE));
  SetControlProperty('ADF_NUMEROCONTACT', 'Plus', GetCodeAuxi('FAC'));

  ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], GetCodeAuxi('FAC'), ValeurI(GetControlText('ADF_NUMEROCONTACT')));
  SetValeursContact('ADF');
end;

SetControlChecked('CP_ADR_IDENT',LaTOB.GetValue('GP_NUMADRESSEFACT') = LaTOB.GetValue('GP_NUMADRESSELIVR'));
if LaTOB.GetValue('GP_NUMADRESSEFACT') = LaTOB.GetValue('GP_NUMADRESSELIVR') then
  SetAdresseFacturationEnabled(False);
PutEcranTobPieceAdresseLivraison( TobAdresses.Detail[0] );
PutEcranTobPieceAdresseFacturation( TobAdresses.Detail[1] );

if Not(ctxGCAFF in V_PGI.PGIContexte) then SetControlVisible ('GBRESSOURCES', False);

{$IFDEF MODE}
SetControlVisible('GP_RESSOURCE', False) ;
SetControlVisible('TGP_RESSOURCE', False) ;
if (LaTOB<>Nil) and (GetInfoParPiece(LaTOB.GetValue('GP_NATUREPIECEG'),'GPP_MAJINFOTIERS')='-') then
  BEGIN
  SetControlEnabled('GBLIVRE', False) ;
  SetControlProperty('GP_TIERSLIVRE', 'Color', clBtnFace) ;
  SetControlProperty('ADL_LIBTIERS', 'Color', clBtnFace) ;
  SetControlProperty('ADL_LIBTIERS2', 'Color', clBtnFace) ;
  SetControlProperty('ADL_ADRESSE1', 'Color', clBtnFace) ;
  SetControlProperty('ADL_ADRESSE2', 'Color', clBtnFace) ;
  SetControlProperty('ADL_ADRESSE3', 'Color', clBtnFace) ;
  SetControlProperty('ADL_CODEPOSTAL', 'Color', clBtnFace) ;
  SetControlProperty('ADL_VILLE', 'Color', clBtnFace) ;
  SetControlProperty('ADL_PAYS', 'Color', clBtnFace) ;
  SetControlProperty('ADL_INCOTERM'    , 'Color', clBtnFace);   
  SetControlProperty('ADL_MODEEXP'     , 'Color', clBtnFace);   
  SetControlProperty('ADL_LIEUDISPO'   , 'Color', clBtnFace);   
  SetControlProperty('ADL_NIF'         , 'Color', clBtnFace);   
  SetControlProperty('ADL_EAN'         , 'Color', clBtnFace);   
  SetControlProperty('ADL_REGION'      , 'Color', clBtnFace);   
  SetControlProperty('ADL_NUMEROCONTACT'  ,'Color', clBtnFace); 
  SetControlProperty('ADL_CONTACT'        ,'Color', clBtnFace); 
  SetControlProperty('ADL_CONTACTPRENOM'  ,'Color', clBtnFace); 
  SetControlEnabled('GBFACTURE', False) ;
  SetControlEnabled('CP_ADR_IDENT', False) ;
  SetControlVisible('BRECHERCHE', False) ;
  END ;
{$ENDIF}
//  BBI, correction fiche 10336
if (ctxMode in V_PGI.PGIContexte) then SetControlVisible('GBPIECE',False);
//  BBI, fin correction fiche 10336
{$IFNDEF CCS3}
AppliquerConfidentialite(Ecran,GetInfoParPiece(NatPiece,'GPP_VENTEACHAT'));
{$ENDIF}

if (VH_GC.GCIfDefCEGID) and (latob.getvalue('GP_DOMAINE')='001' ) then
begin
  if Assigned(GetControl('GBDATEPIECE')) then
    SetControlVisible('GBDATEPIECE', true);
end;
end;

procedure TOF_ComplPiece.OnUpdate;
Var VenteAchat, NomChamp : string;
    LibDuChamp : string; // JT - eQualité 10819
begin
// Inherited;
// Contrôle de la ressource saisie à la ligne
  if (GetControlText('GP_RESSOURCE') <> '' ) then
     if Not(LookupValueExist(THEdit(Getcontrol('GP_RESSOURCE')))) then
        BEGIN
        PGIBox('Vous devez renseigner une ressource valide',Ecran.Caption);
        SetFocusControl('GP_RESSOURCE') ;
        TForm(Ecran).ModalResult:=0; Exit ;
        END;

VenteAchat:= LaTob.GetValue('GP_VENTEACHAT') ;
if (VenteAchat= 'VEN') or ((VenteAchat='ACH') and (GetControlText('GP_TIERSLIVRE')<>'')) then
   begin
   if not ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+
           GetControlText('GP_TIERSLIVRE')+'"') then
      begin
      SetActiveTabSheet('GENERAL');
      SetLastError(1, 'GP_TIERSLIVRE');
      Exit ;
    end ;
    if (GetControlText('ADL_LIBTIERS') = '' ) then
    begin
      SetLastError(2, 'ADL_LIBTIERS');
      Exit ;
    end;
    if (GetControlText('ADL_CODEPOSTAL') = '' ) then
    begin
      SetLastError(3, 'ADL_CODEPOSTAL');
      Exit ;
    end;
    if (GetControlText('ADL_VILLE') = '' ) then
    begin
      SetLastError(4, 'ADL_VILLE');
      exit ;
      end ;
   end ;
if (VenteAchat= 'VEN') or ((VenteAchat='ACH') and (GetControlText('GP_TIERSFACTURE')<>'')) then
   begin
   if not ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+
           GetControlText('GP_TIERSFACTURE')+'"') then
      begin
      SetActiveTabSheet('GENERAL');
      SetLastError(1, 'GP_TIERSFACTURE');
      exit ;
      end ;
    if (GetControlText('ADF_LIBTIERS') = '' ) then
    begin
      SetLastError(5, 'ADF_LIBTIERS');
      Exit ;
    end ;
    if (GetControlText('ADF_CODEPOSTAL') = '' ) then
    begin
      SetLastError(6, 'ADF_CODEPOSTAL');
      Exit ;
    end;
    if (GetControlText('ADF_VILLE') = '' ) then
    begin
      SetLastError(7, 'ADF_VILLE');
      Exit ;
    end;
  end;

{$IFNDEF CCS3}
NomChamp:=VerifierChampsObligatoires(Ecran,VenteAchat);
   if NomChamp<>'' then
      begin
      NomChamp:=ReadTokenSt(NomChamp);
      SetFocusControl(NomChamp) ;
      LibDuChamp := LibTableLibreCliFour(VenteAchat, NomChamp); // JT - eQualité 10819
      PGIBox('La saisie du champs suivant est obligatoire : '+LibDuChamp,Ecran.Caption);
      TForm(Ecran).ModalResult:=0;
      Exit ;
      end;
{$ENDIF}
if GetParamSoc('SO_GCPIECEADRESSE') then
BEGIN
  TOBAdresses.Detail[0].PutValue('GPA_LIBELLE'        ,GetControlText('ADL_LIBTIERS')     );
  TOBAdresses.Detail[0].PutValue('GPA_LIBELLE2'       ,GetControlText('ADL_LIBTIERS2')    );
  TOBAdresses.Detail[0].PutValue('GPA_ADRESSE1'       ,GetControlText('ADL_ADRESSE1')     );
  TOBAdresses.Detail[0].PutValue('GPA_ADRESSE2'       ,GetControlText('ADL_ADRESSE2')     );
  TOBAdresses.Detail[0].PutValue('GPA_ADRESSE3'       ,GetControlText('ADL_ADRESSE3')     );
  TOBAdresses.Detail[0].PutValue('GPA_CODEPOSTAL'     ,GetControlText('ADL_CODEPOSTAL')   );
  TOBAdresses.Detail[0].PutValue('GPA_VILLE'          ,GetControlText('ADL_VILLE')        );
  TOBAdresses.Detail[0].PutValue('GPA_PAYS'           ,GetControlText('ADL_PAYS')         );
  TOBAdresses.Detail[0].PutValue('GPA_INCOTERM'       ,GetControlText('ADL_INCOTERM')     ); 
  TOBAdresses.Detail[0].PutValue('GPA_MODEEXP'        ,GetControlText('ADL_MODEEXP')      ); 
  TOBAdresses.Detail[0].PutValue('GPA_LIEUDISPO'      ,GetControlText('ADL_LIEUDISPO')    ); 
  TOBAdresses.Detail[0].PutValue('GPA_NIF'            ,GetControlText('ADL_NIF')          ); 
  TOBAdresses.Detail[0].PutValue('GPA_EAN'            ,GetControlText('ADL_EAN')          ); 
  TOBAdresses.Detail[0].PutValue('GPA_REGION'         ,GetControlText('ADL_REGION')       ); 
  TOBAdresses.Detail[0].PutValue('GPA_NUMEROCONTACT'  ,GetControlText('ADL_NUMEROCONTACT')); 
END
else
BEGIN
  TOBAdresses.Detail[0].PutValue('ADR_LIBELLE'      ,GetControlText('ADL_LIBTIERS')     );
  TOBAdresses.Detail[0].PutValue('ADR_LIBELLE2'     ,GetControlText('ADL_LIBTIERS2')    );
  TOBAdresses.Detail[0].PutValue('ADR_ADRESSE1'     ,GetControlText('ADL_ADRESSE1')     );
  TOBAdresses.Detail[0].PutValue('ADR_ADRESSE2'     ,GetControlText('ADL_ADRESSE2')     );
  TOBAdresses.Detail[0].PutValue('ADR_ADRESSE3'     ,GetControlText('ADL_ADRESSE3')     );
  TOBAdresses.Detail[0].PutValue('ADR_CODEPOSTAL'   ,GetControlText('ADL_CODEPOSTAL')   );
  TOBAdresses.Detail[0].PutValue('ADR_VILLE'        ,GetControlText('ADL_VILLE')        );
  TOBAdresses.Detail[0].PutValue('ADR_PAYS'         ,GetControlText('ADL_PAYS')         );
  TOBAdresses.Detail[0].PutValue('ADR_INCOTERM'     ,GetControlText('ADL_INCOTERM')     ); 
  TOBAdresses.Detail[0].PutValue('ADR_MODEEXP'      ,GetControlText('ADL_MODEEXP')      ); 
  TOBAdresses.Detail[0].PutValue('ADR_LIEUDISPO'    ,GetControlText('ADL_LIEUDISPO')    ); 
  TOBAdresses.Detail[0].PutValue('ADR_NIF'          ,GetControlText('ADL_NIF')          ); 
  TOBAdresses.Detail[0].PutValue('ADR_EAN'          ,GetControlText('ADL_EAN')          ); 
  TOBAdresses.Detail[0].PutValue('ADR_REGION'       ,GetControlText('ADL_REGION')       ); 
  TOBAdresses.Detail[0].PutValue('ADR_NUMEROCONTACT',GetControlText('ADL_NUMEROCONTACT')); 
END ;
LaTOB.GetEcran(TForm(Ecran));
if TCheckBox(GetControl('CP_ADR_IDENT')).Checked then
begin
  LaTOB.PutValue('GP_NUMADRESSEFACT',LaTOB.GetValue('GP_NUMADRESSELIVR')) ;
  TOBAdresses.Detail[1].Dupliquer(TOBAdresses.Detail[0],False,True) ;
  if GetParamSoc('SO_GCPIECEADRESSE') then TOBAdresses.Detail[1].PutValue('GPA_TYPEPIECEADR','002');
end
else
begin
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    TOBAdresses.Detail[1].PutValue('GPA_LIBELLE'    ,GetControlText('ADF_LIBTIERS'));
    TOBAdresses.Detail[1].PutValue('GPA_LIBELLE2'   ,GetControlText('ADF_LIBTIERS2'));
    TOBAdresses.Detail[1].PutValue('GPA_ADRESSE1'   ,GetControlText('ADF_ADRESSE1'));
    TOBAdresses.Detail[1].PutValue('GPA_ADRESSE2'   ,GetControlText('ADF_ADRESSE2'));
    TOBAdresses.Detail[1].PutValue('GPA_ADRESSE3'   ,GetControlText('ADF_ADRESSE3'));
    TOBAdresses.Detail[1].PutValue('GPA_CODEPOSTAL' ,GetControlText('ADF_CODEPOSTAL'));
    TOBAdresses.Detail[1].PutValue('GPA_VILLE'      ,GetControlText('ADF_VILLE'));
    TOBAdresses.Detail[1].PutValue('GPA_PAYS'       ,GetControlText('ADF_PAYS'));
    TOBAdresses.Detail[1].PutValue('GPA_INCOTERM'   ,GetControlText('ADF_INCOTERM')  ); 
    TOBAdresses.Detail[1].PutValue('GPA_MODEEXP'    ,GetControlText('ADF_MODEEXP')   ); 
    TOBAdresses.Detail[1].PutValue('GPA_LIEUDISPO'  ,GetControlText('ADF_LIEUDISPO') ); 
    TOBAdresses.Detail[1].PutValue('GPA_NIF'        ,GetControlText('ADF_NIF')       ); 
    TOBAdresses.Detail[1].PutValue('GPA_EAN'        ,GetControlText('ADF_EAN')       ); 
    TOBAdresses.Detail[1].PutValue('GPA_REGION'     ,GetControlText('ADF_REGION')    ); 
    TOBAdresses.Detail[1].PutValue('GPA_NUMEROCONTACT'  ,GetControlText('ADF_NUMEROCONTACT') ); 
    if LaTOB.GetValue('GP_NUMADRESSEFACT')=LaTOB.GetValue('GP_NUMADRESSELIVR') then LaTOB.PutValue('GP_NUMADRESSEFACT',+2) ;
    end
    else
    begin
      TOBAdresses.Detail[1].PutValue('ADR_LIBELLE'      ,GetControlText('ADF_LIBTIERS')  );
      TOBAdresses.Detail[1].PutValue('ADR_LIBELLE2'     ,GetControlText('ADF_LIBTIERS2') );
      TOBAdresses.Detail[1].PutValue('ADR_ADRESSE1'     ,GetControlText('ADF_ADRESSE1')  );
      TOBAdresses.Detail[1].PutValue('ADR_ADRESSE2'     ,GetControlText('ADF_ADRESSE2')  );
      TOBAdresses.Detail[1].PutValue('ADR_ADRESSE3'     ,GetControlText('ADF_ADRESSE3')  );
      TOBAdresses.Detail[1].PutValue('ADR_CODEPOSTAL'   ,GetControlText('ADF_CODEPOSTAL'));
      TOBAdresses.Detail[1].PutValue('ADR_VILLE'        ,GetControlText('ADF_VILLE')     );
      TOBAdresses.Detail[1].PutValue('ADR_PAYS'         ,GetControlText('ADF_PAYS')      );
      TOBAdresses.Detail[1].PutValue('ADR_INCOTERM'     ,GetControlText('ADF_INCOTERM')  ); 
      TOBAdresses.Detail[1].PutValue('ADR_MODEEXP'      ,GetControlText('ADF_MODEEXP')   ); 
      TOBAdresses.Detail[1].PutValue('ADR_LIEUDISPO'    ,GetControlText('ADF_LIEUDISPO') ); 
      TOBAdresses.Detail[1].PutValue('ADR_NIF'          ,GetControlText('ADF_NIF')       ); 
      TOBAdresses.Detail[1].PutValue('ADR_EAN'          ,GetControlText('ADF_EAN')       ); 
      TOBAdresses.Detail[1].PutValue('ADR_REGION'       ,GetControlText('ADF_REGION')    ); 
      TOBAdresses.Detail[1].PutValue('ADR_NUMEROCONTACT',GetControlText('ADF_NUMEROCONTACT') ); 
      if LaTOB.GetValue('GP_NUMADRESSEFACT')=LaTOB.GetValue('GP_NUMADRESSELIVR') then LaTOB.PutValue('GP_NUMADRESSEFACT',-2) ;
    end;
  end;
if (VenteAchat='ACH') and (GetControlText('GP_TIERSLIVRE')='') then LaTOB.PutValue('GP_NUMADRESSELIVR',0);
if (VenteAchat='ACH') and (GetControlText('GP_TIERSFACTURE')='') then LaTOB.PutValue('GP_NUMADRESSEFACT',0);
TheTob:=LaTOB ;
end;

procedure TOF_ComplPiece_InfoTiers( Parms: array of variant; nb: integer ) ;
Var F : TForm;
    TOBTiers : TOB ;
    i : integer ;
    CodeTiers, CodeAuxi, Prefixe : string;
    BoxLivre, BoxFacture : TGroupBox ;
begin
F := TForm(Longint(Parms[0]));
if F.Name <> 'GCCOMPLPIECE' then exit;
CodeTiers:=string(Parms[1]);
Prefixe:=string(Parms[2]);
if (Prefixe='ADL') then
   begin
   if ComplPieceCodeLivre=CodeTiers then exit
   else ComplPieceCodeLivre:=CodeTiers;
   end;
if Prefixe='ADF' then
   begin
   if ComplPieceCodeFacture=CodeTiers then exit
   else ComplPieceCodeFacture:=CodeTiers;
   end;
if bAdrOk then begin bAdrOk := false; exit; end;
CodeAuxi:=TiersAuxiliaire(CodeTiers,false);
if CodeAuxi = '' then
  begin
  if Prefixe='ADL' then
     begin
     BoxLivre:=TGroupBox(F.FindComponent('GBLIVRE'));
     For i:=0 to BoxLivre.ControlCount-1 do
         if (BoxLivre.Controls[i] is THEdit) and
            (THEdit(BoxLivre.Controls[i]).Name<>'GP_TIERSLIVRE') then
                                     THEdit(BoxLivre.Controls[i]).Text:='';
     end else if Prefixe='ADF' then
     begin
     BoxFacture:=TGroupBox(F.FindComponent('GBFACTURE'));
     For i:=0 to BoxFacture.ControlCount-1 do
         if BoxFacture.Controls[i] is THEdit and
                 (THEdit(BoxFacture.Controls[i]).Name<>'GP_TIERSFACTURE') then
                                   THEdit(BoxFacture.Controls[i]).Text:='';
     end ;
  exit ;
  end ;
TOBTiers := TOB.Create('TIERS', nil, -1);
TOBTiers.SelectDB('"'+CodeAuxi+'"',Nil);
THEdit(F.FindComponent(Prefixe+'_LIBTIERS')).Text := TOBTiers.GetValue('T_LIBELLE');
THEdit(F.FindComponent(Prefixe+'_LIBTIERS2')).Text := TOBTiers.GetValue('T_PRENOM');
THEdit(F.FindComponent(Prefixe+'_ADRESSE1')).Text := TOBTiers.GetValue('T_ADRESSE1');
THEdit(F.FindComponent(Prefixe+'_ADRESSE2')).Text := TOBTiers.GetValue('T_ADRESSE2');
THEdit(F.FindComponent(Prefixe+'_ADRESSE3')).Text := TOBTiers.GetValue('T_ADRESSE3');
THEdit(F.FindComponent(Prefixe+'_CODEPOSTAL')).Text := TOBTiers.GetValue('T_CODEPOSTAL');
THEdit(F.FindComponent(Prefixe+'_VILLE')).Text := TOBTiers.GetValue('T_VILLE');
THEdit(F.FindComponent(Prefixe+'_PAYS')).Text := TOBTiers.GetValue('T_PAYS');
TOBTiers.Free;
end;

procedure TOF_ComplPiece_POPZ( Parms: array of variant; nb: integer ) ;
var
    F : TForm;
    Option : Integer;
    CodeTiers, Valeur, Params, Range : string;
    TOBAdresses : TOB;
    ValuesContact : MyArrayValue; 
    CodeAuxi : String;            

begin
if VH_GC.GCIfDefCEGID then
  Exit ; // à revoir plus tard
F := TForm(Longint(Parms[0]));
Option := Integer(Parms[1]);
TOBAdresses := TOB.Create('ADRESSES', nil, -1);
case Option of
  1 : begin
        CodeTiers := THEdit(F.FindComponent('GP_TIERSLIVRE')).Text;
        CodeAuxi:=TiersAuxiliaire(CodeTiers,false); 
        Range :='ADR_REFCODE=' + CodeTiers + ';ADR_LIVR=X';
        Params := 'TYPEADRESSE=TIE;NATUREAUXI=CLI'
                +';YTC_TIERSLIVRE=' + CodeTiers
                +';CLI=' + CodeAuxi
                +';TIERS=' + CodeTiers
                +';ACTION=MODIFICATION';

        Valeur := AglLanceFiche('GC', 'GCADRESSES', Range, '', Params);
        if Valeur = '' then begin TOBAdresses.Free; exit; end;
        bAdrOk := THEdit(F.FindComponent('GP_TIERSLIVRE')).Focused;
        TOBAdresses.PutValue('ADR_NUMEROADRESSE', StrToInt (Valeur));
        TOBAdresses.PutValue('ADR_TYPEADRESSE', 'TIE');
        TOBAdresses.LoadDB();
        THEdit(F.FindComponent('GP_NUMADRESSELIVR')  ).Text := TOBAdresses.GetValue('ADR_NUMEROADRESSE');
        THEdit(F.FindComponent('GP_TIERSLIVRE') ).Text := TOBAdresses.GetValue('ADR_REFCODE');
        if not ExisteSQL('SELECT ADR_FACT FROM ADRESSES WHERE ADR_REFCODE = "'+CodeTiers+'" AND ADR_NUMEROADRESSE = '+THEdit(F.FindComponent('GP_NUMADRESSELIVR')).Text + ' AND ADR_FACT="X"') then
          TCheckBox(F.FindComponent('CP_ADR_IDENT')).State := CbUnChecked;

//      THEdit(F.FindComponent('GP_NADRESSELIV')).Text := TOBAdresses.GetValue('ADR_NADRESSE'); 
        THEdit(F.FindComponent('ADL_LIBTIERS')  ).Text := TOBAdresses.GetValue('ADR_LIBELLE');
        THEdit(F.FindComponent('ADL_LIBTIERS2') ).Text := TOBAdresses.GetValue('ADR_LIBELLE2');
        THEdit(F.FindComponent('ADL_ADRESSE1')  ).Text := TOBAdresses.GetValue('ADR_ADRESSE1');
        THEdit(F.FindComponent('ADL_ADRESSE2')  ).Text := TOBAdresses.GetValue('ADR_ADRESSE2');
        THEdit(F.FindComponent('ADL_ADRESSE3')  ).Text := TOBAdresses.GetValue('ADR_ADRESSE3');
        THEdit(F.FindComponent('ADL_CODEPOSTAL')).Text := TOBAdresses.GetValue('ADR_CODEPOSTAL');
        THEdit(F.FindComponent('ADL_VILLE')     ).Text := TOBAdresses.GetValue('ADR_VILLE');
        thValcomboBox(F.FindComponent('ADL_PAYS')     ).Value := TOBAdresses.GetValue('ADR_PAYS');
        THEdit(F.FindComponent('ADL_INCOTERM') ).Text := TOBAdresses.GetValue('ADR_INCOTERM'); 
        thValcomboBox(F.FindComponent('ADL_MODEEXP')  ).Value := TOBAdresses.GetValue('ADR_MODEEXP');  
        thValcomboBox(F.FindComponent('ADL_LIEUDISPO')).Value := TOBAdresses.GetValue('ADR_LIEUDISPO');
        THEdit(F.FindComponent('ADL_NIF')).Text := TOBAdresses.GetValue('ADR_NIF');                    
        THEdit(F.FindComponent('ADL_EAN')).Text := TOBAdresses.GetValue('ADR_EAN');                    
        THValComboBox(F.FindComponent('ADL_REGION')).Value := TOBAdresses.GetValue('ADR_REGION');      
        THEdit(F.FindComponent('ADL_NUMEROCONTACT')).Text := TOBAdresses.GetValue('ADR_NUMEROCONTACT');                
        ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], CodeAuxi, TobAdresses.GetValue('ADR_NUMEROCONTACT')); 
        if ValuesContact <> nil then                                                                                   
        begin                                                                                                          
          THEdit(F.FindComponent('ADL_CONTACT')).Text := ValuesContact[0] ;                                            
          THEdit(F.FindComponent('ADL_CONTACTPRENOM')).Text := ValuesContact[1];                                       
        end                                                                                                            
        else                                                                                                           
        begin                                                                                                          
          THEdit(F.FindComponent('ADL_CONTACT')).Text := '' ;                                                          
          THEdit(F.FindComponent('ADL_CONTACTPRENOM')).Text := '';                                                     
        end;                                                                                                           
        if TCheckBox(F.FindComponent('CP_ADR_IDENT')).Checked then
        begin
          THEdit(F.FindComponent('GP_TIERSFACTURE') ).Text := TOBAdresses.GetValue('ADR_REFCODE');
//        THEdit(F.FindComponent('GP_NADRESSEFAC')  ).Text := TOBAdresses.GetValue('ADR_NADRESSE');
          THEdit(F.FindComponent('ADF_LIBTIERS')    ).Text := TOBAdresses.GetValue('ADR_LIBELLE');
          THEdit(F.FindComponent('ADF_LIBTIERS2')   ).Text := TOBAdresses.GetValue('ADR_LIBELLE2');
          THEdit(F.FindComponent('ADF_ADRESSE1')    ).Text := TOBAdresses.GetValue('ADR_ADRESSE1');
          THEdit(F.FindComponent('ADF_ADRESSE2')    ).Text := TOBAdresses.GetValue('ADR_ADRESSE2');
          THEdit(F.FindComponent('ADF_ADRESSE3')    ).Text := TOBAdresses.GetValue('ADR_ADRESSE3');
          THEdit(F.FindComponent('ADF_CODEPOSTAL')  ).Text := TOBAdresses.GetValue('ADR_CODEPOSTAL');
          THEdit(F.FindComponent('ADF_VILLE')       ).Text := TOBAdresses.GetValue('ADR_VILLE');
          thValcomboBox(F.FindComponent('ADF_PAYS')     ).Value := TOBAdresses.GetValue('ADR_PAYS');
          THEdit(F.FindComponent('ADF_INCOTERM') ).Text := TOBAdresses.GetValue('ADR_INCOTERM'); 
          thValcomboBox(F.FindComponent('ADF_MODEEXP')  ).Value := TOBAdresses.GetValue('ADR_MODEEXP');  
          thValcomboBox(F.FindComponent('ADF_LIEUDISPO')).Value := TOBAdresses.GetValue('ADR_LIEUDISPO');
          THEdit(F.FindComponent('ADF_NIF')         ).Text := TOBAdresses.GetValue('ADR_NIF');           
          THEdit(F.FindComponent('ADF_EAN')         ).Text := TOBAdresses.GetValue('ADR_EAN');           
          THValComboBox(F.FindComponent('ADF_REGION')).Value := TOBAdresses.GetValue('ADR_REGION');      
          THEdit(F.FindComponent('ADF_NUMEROCONTACT')).Text := TOBAdresses.GetValue('ADR_NUMEROCONTACT');                 
          ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], CodeAuxi, TobAdresses.GetValue('ADR_NUMEROCONTACT'));  
          if ValuesContact <> nil then                                                                                    
          begin                                                                                                           
            THEdit(F.FindComponent('ADF_CONTACT')).Text := ValuesContact[0];                                              
            THEdit(F.FindComponent('ADF_CONTACTPRENOM')).Text := ValuesContact[1];                                        
          end                                                                                                             
          else                                                                                                            
          begin                                                                                                           
            THEdit(F.FindComponent('ADF_CONTACT')).Text := '';                                                            
            THEdit(F.FindComponent('ADF_CONTACTPRENOM')).Text := '';                                                      
          end;                                                                                                            
        end;
      end;
  2 : begin
        CodeTiers := THEdit(F.FindComponent('GP_TIERSFACTURE')).Text;
        CodeAuxi  := TiersAuxiliaire(CodeTiers,false); 
        Range :='ADR_REFCODE=' + CodeTiers + ';ADR_FACT=X';
        Params := 'TYPEADRESSE=TIE;NATUREAUXI=CLI'
                +';YTC_TIERSLIVRE=' + CodeTiers
                +';CLI=' + CodeAuxi
                +';TIERS=' + CodeTiers
                +';ACTION=MODIFICATION';

        Valeur := AglLanceFiche('GC', 'GCADRESSES', Range, '', Params);
//        Valeur := AGLLanceFiche('GC','GCADRESSES_MUL', 'ADR_REFCODE=' + CodeTiers + '; ADR_FACT=X', '', Select);
        if Valeur = '' then begin TOBAdresses.Free; exit; end;
        bAdrOk := THEdit(F.FindComponent('GP_TIERSFACTURE')).Focused;
        TOBAdresses.PutValue('ADR_NUMEROADRESSE', StrToInt (Valeur));
        TOBAdresses.PutValue('ADR_TYPEADRESSE', 'TIE');
        TOBAdresses.LoadDB();
        THEdit(F.FindComponent('GP_NUMADRESSEFACT')  ).Text := TOBAdresses.GetValue('ADR_NUMEROADRESSE');
        THEdit(F.FindComponent('GP_TIERSFACTURE')  ).Text := TOBAdresses.GetValue('ADR_REFCODE');
//      THEdit(F.FindComponent('GP_NADRESSEFAC')   ).Text := TOBAdresses.GetValue('ADR_NADRESSE');    
        THEdit(F.FindComponent('ADF_LIBTIERS')     ).Text := TOBAdresses.GetValue('ADR_LIBELLE');
        THEdit(F.FindComponent('ADF_LIBTIERS2')    ).Text := TOBAdresses.GetValue('ADR_LIBELLE2');
        THEdit(F.FindComponent('ADF_ADRESSE1')     ).Text := TOBAdresses.GetValue('ADR_ADRESSE1');
        THEdit(F.FindComponent('ADF_ADRESSE2')     ).Text := TOBAdresses.GetValue('ADR_ADRESSE2');
        THEdit(F.FindComponent('ADF_ADRESSE3')     ).Text := TOBAdresses.GetValue('ADR_ADRESSE3');
        THEdit(F.FindComponent('ADF_CODEPOSTAL')   ).Text := TOBAdresses.GetValue('ADR_CODEPOSTAL');
        THEdit(F.FindComponent('ADF_VILLE')        ).Text := TOBAdresses.GetValue('ADR_VILLE');
        thValcomboBox(F.FindComponent('ADF_PAYS')     ).Value := TOBAdresses.GetValue('ADR_PAYS');
        THEdit(F.FindComponent('ADF_INCOTERM') ).Text := TOBAdresses.GetValue('ADR_INCOTERM'); 
        thValcomboBox(F.FindComponent('ADF_MODEEXP')  ).Value := TOBAdresses.GetValue('ADR_MODEEXP');  
        thValcomboBox(F.FindComponent('ADF_LIEUDISPO')).Value := TOBAdresses.GetValue('ADR_LIEUDISPO');
        THEdit(F.FindComponent('ADF_NIF')          ).Text := TOBAdresses.GetValue('ADR_NIF');          
        THEdit(F.FindComponent('ADF_EAN')          ).Text := TOBAdresses.GetValue('ADR_EAN');
        THValComboBox(F.FindComponent('ADF_REGION')).Value := TOBAdresses.GetValue('ADR_REGION');
        THEdit(F.FindComponent('ADF_NUMEROCONTACT')).Text := TOBAdresses.GetValue('ADR_NUMEROCONTACT');
        ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], CodeAuxi, TobAdresses.GetValue('ADR_NUMEROCONTACT'));  
        if ValuesContact <> nil then                                                                                    
        begin                                                                                                         
          THEdit(F.FindComponent('ADF_CONTACT')).Text := ValuesContact[0];                                            
          THEdit(F.FindComponent('ADF_CONTACTPRENOM')).Text := ValuesContact[1];                                      
        end                                                                                                           
        else                                                                                                          
        begin                                                                                                         
          THEdit(F.FindComponent('ADF_CONTACT')).Text := '';                                                          
          THEdit(F.FindComponent('ADF_CONTACTPRENOM')).Text := '';                                                    
        end;
      end;
    end;
TOBAdresses.Free;
end;

/////////////////////////////////////////////////////////////////////////////
procedure TOF_SaisCond.OnCancel;
begin
    Inherited;
TOBCond.Free;
end;

procedure TOF_SaisCond.OnNew;
begin
    Inherited;
end;

procedure TOF_SaisCond.OnDelete;
begin
    Inherited;
end;

procedure TOF_SaisCond.OnLoad;
var F : TForm;
    CodeCond : string;
    TOBC : TOB;
begin
    Inherited;
F := TForm (Ecran);
//AgentAnimateSt ('Suggest');
TOBCond := TOB.Create('_', nil, -1);
TOBCond.Dupliquer(LaTOB, True, True);
CodeCond := LaTOB.GetValue ('GCO_CODECOND');
TOBC:=TOBCond.FindFirst(['GCO_CODECOND'],[CodeCond],False); if TOBC = nil then exit;
SetControlText('GCO_COEFCONV', TOBC.GetValue('GCO_COEFCONV'));
SetControlText('GCO_ARRONDIINF', TOBC.GetValue('GCO_ARRONDIINF'));
SetControlText('GCO_UNITEAFFICHE', TOBC.GetValue('GCO_UNITEAFFICHE'));
SetControlText('GCO_NBARTICLE', TOBC.GetValue('GCO_NBARTICLE')) ;
SetControlText('GCO_DECOND', TOBC.GetValue('GCO_DECOND')) ;
CalcQte (F, 'QTEARTICLE') ;
end;

procedure TOF_SaisCond.OnUpdate;
begin
    Inherited;
if (LaTOB <> Nil) then TheTob := LaTOB ;
TOBCond.Free;
end;

procedure TOF_SaisCond_CalcQte (Parms: array of variant; nb: integer) ;
var F : TForm;
    QteRef : string;
BEGIN
F := TForm (Longint (Parms[0]));
if F.Name <> 'GCSAISCOND' then exit;
QteRef := String (Parms[1]);
CalcQte (F, QteRef);
END;

procedure CalcQte (F : TForm; QteRef : String) ;
var QteArt, QteUnit, Coef : Double;
    Arrondi : Boolean;
    St : string;
BEGIN
St := THEdit(F.FindComponent('QTEARTICLE')).Text ;
if St = '' then QteArt := 0 else QteArt := StrToFloat (St) ;
St := THEdit(F.FindComponent('QTEUNITE')).Text;
if St = '' then QteUnit := 0 else QteUnit := StrToFloat (St) ;
St := THEdit(F.FindComponent('GCO_COEFCONV')).Text;
if St = '' then Coef := 0 else Coef := StrToFloat (St) ;
Arrondi := TCheckBox (F.FindComponent('GCO_ARRONDIINF')).Checked;
if coef = 0 then exit;
// Qté en unité de saisie = Qté produit * coef de conversion
if QteRef = 'QTEARTICLE' then
    BEGIN
    QteUnit := QteArt * Coef;
    THEdit(F.FindComponent('QTEUNITE')).Text := FloatToStr (QteUnit);
    END else
    BEGIN
    QteArt := QteUnit / Coef;
    If QteArt <> Trunc (QteArt) then
        if Arrondi then
            QteArt := Trunc (QteArt) else
            QteArt := Trunc (QteArt) + 1;
    AdapteCond (F, QteArt);
    END;
END;

procedure TOF_SaisCond_AdapteCond (Parms: array of variant; nb: integer) ;
var F : TForm;
    QteArt : double;
BEGIN
F := TForm (Longint (Parms[0]));
if F.Name <> 'GCSAISCOND' then exit;
if String (Parms[1]) = '' then QteArt := 0 else QteArt := StrToFloat (Parms[1]);
AdapteCond (F, QteArt);
END;

Procedure AdapteCond (F : TForm; QteArt : double) ;
Var St : string ;
    NbArt, PEntSup : Double ;
BEGIN
St:= string (THEdit(F.FindComponent('GCO_NBARTICLE')).Text) ;
if St = '' then NbArt := 1 else NbArt := StrToFloat (St) ;
if NbArt <> 0 then
    BEGIN
    if string (THEdit(F.FindComponent('GCO_DECOND')).Text) <> 'X' then
        BEGIN
        PEntSup := QteArt / NbArt ;
        if Trunc (PEntSup) <> PEntSup then
            BEGIN
            PEntSup := Trunc(PEntSup) + 1 ;
            QteArt := PentSup * NbArt ;
            END;
        END ;
    END ;
THEdit(F.FindComponent('QTEARTICLE')).Text := FloatToStr (QteArt);
END;

Function Mul_DeterminePieceAutorise (Parms: array of variant; nb: integer) : variant ;
var F : TForm;
    NatPiece, VenteAchat, St_OR, St_Where : string;
    QQ : TQuery ;
    TOBP, TOBPL : TOB;
    i_ind : integer ;
BEGIN
Result := '';
F := TForm (Longint (Parms[0]));
NatPiece := String (Parms[1]) ;
VenteAchat := '';
QQ:=OpenSQL('SELECT GPP_VENTEACHAT FROM PARPIECE ' +
                'WHERE GPP_NATUREPIECEG="' + NatPiece + '"', True) ;
if (not QQ.Eof) then VenteAchat := QQ.Fields[0].AsString ;
Ferme (QQ);
if VenteAchat <> '' then
  THValComboBox(F.FindComponent('GP_NATUREPIECEG')).Plus:='AND GPP_VENTEACHAT="'+VenteAchat+'"';
if (F.Name = 'BTTRANSPIECE_MUL') or (F.Name = 'GCTRANSPIECE_MUL') or (F.Name = 'GCTRANSACH_MUL') then
    BEGIN
    St_OR := '' ;
    St_Where := 'GPP_NATUREPIECEG<>"' + NatPiece + '"' ;
    St_Where := St_Where +  ' AND GPP_NATUREPIECEG<>"AFF" AND GPP_NATUREPIECEG<>"PAF"';
    QQ:=OpenSQL('SELECT GPP_NATUREPIECEG FROM PARPIECE ' +
                'WHERE GPP_NATURESUIVANTE like "%' + NatPiece + '%" AND ' + St_Where, True) ;
    TOBP := TOB.Create ('', nil, -1);
    TOBP.LoadDetailDB ('', '', '', QQ, False) ;
    Ferme (QQ) ;
    for i_ind := 0 to TOBP.Detail.count - 1 do
        BEGIN
        TOBPL := TOBP.Detail[i_ind] ;
        Result := Result + St_OR + 'GP_NATUREPIECEG="' + TOBPL.GetValue ('GPP_NATUREPIECEG') + '"';
        St_OR := ' OR ' ;
        END;
    TOBP.free ;
    END;
    if (F.Name = 'GCDUPLICPIECE_MUL') or (F.Name = 'BTDUPLICPIECE_MUL') then
    BEGIN
    if VenteAchat <> '' then
        BEGIN
        St_OR := '' ;
        QQ:=OpenSQL('SELECT GPP_NATUREPIECEG FROM PARPIECE ' +
                'WHERE GPP_VENTEACHAT="' + VenteAchat + '"', True) ;
        TOBP := TOB.Create ('', nil, -1);
        TOBP.LoadDetailDB ('', '', '', QQ, False) ;
        Ferme (QQ) ;
        for i_ind := 0 to TOBP.Detail.count - 1 do
            BEGIN
            TOBPL := TOBP.Detail[i_ind] ;
            Result := Result + St_OR + 'GP_NATUREPIECEG="' + TOBPL.GetValue ('GPP_NATUREPIECEG') + '"';
            St_OR := ' OR ' ;
            END;
        TOBP.free ;
        END;
    END;
END;

/////////////////////////////////////////////////////////////////////////////

procedure InitTOFPiece ();
begin
RegisterAglProc('ComplPiece_POPZ', True , 1, TOF_ComplPiece_POPZ);
RegisterAglProc('ComplPiece_InfoTiers', True , 2, TOF_ComplPiece_InfoTiers);
RegisterAglProc('SaisCond_CalcQte', True , 1, TOF_SaisCond_CalcQte);
RegisterAglFunc('DeterminePieceAutorise', True , 1, Mul_DeterminePieceAutorise);
RegisterAglProc('SaisCond_AdapteCond', True , 1, TOF_SaisCond_CalcQte); // JT eQualité 10998
end;

procedure TOF_ComplPiece.NUMEROCONTACTOnEnter(Sender: TObject);
begin
   { Sauve le contact en entrée }
  if TControl(Sender).Name = 'ADL_NUMEROCONTACT' then
    NumeroContactEnter := GetControlText('ADL_NUMEROCONTACT')
  else
    NumeroContactEnter := GetControlText('ADF_NUMEROCONTACT')
  ;
end;

procedure TOF_ComplPiece.NUMEROCONTACTOnExit(Sender: TObject);
Var
  Prefixe : string;
begin
  if TControl(Sender).Name = 'ADL_NUMEROCONTACT' then
  begin
    prefixe:= 'ADL';
  end
  else
  begin
    prefixe := 'ADF';
  end;

  { Si le n° du contact à changé  }
  if NumeroContactEnter <> GetControlText(prefixe +'_NUMEROCONTACT') then
  begin
    if (LookupValueExist(GetControl(prefixe + '_NUMEROCONTACT'))) then
    begin
      ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], GetCodeAuxi(prefixe), ValeurI(GetControlText(prefixe + '_NUMEROCONTACT')));
      if ValuesContact <> nil then
      begin
        SetControlText(prefixe + '_CONTACT', ValuesContact[0]);
        SetControlText(prefixe + '_CONTACTPRENOM', ValuesContact[1]);
      end;
    end
    else
    begin
      SetControlText(prefixe + '_NUMEROCONTACT', '0');
      SetControlText(prefixe + '_CONTACT', '');
      SetControlText(prefixe + '_CONTACTPRENOM', '');
    end;

    if (TControl(Sender).Name = 'ADL_NUMEROCONTACT') and TCheckBox(GetControl('CP_ADR_IDENT')).Checked then
    begin
      SetControlText('ADF_NUMEROCONTACT',GetControlText('ADL_NUMEROCONTACT'));
      ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], GetCodeAuxi(Prefixe), ValeurI(GetControlText('ADF_NUMEROCONTACT')));
      SetValeursContact('ADF');
    end;
  end;
end;

procedure TOF_ComplPiece.SetValeursContact(Prefixe: string);
begin
  if ValuesContact <> nil then
  begin
    SetControlText(Prefixe + '_CONTACT'        ,ValuesContact[0]     );
    SetControlText(Prefixe + '_CONTACTPRENOM'  ,ValuesContact[1]     );
  end
  else
  begin
    SetControlText(Prefixe + '_CONTACT'        ,''     );
    SetControlText(Prefixe + '_CONTACTPRENOM'  ,''     );
  end
end;

Function TOF_ComplPiece.GetCodeAuxi(Prefixe: string) : string;
var
  Suffixe: string;
begin
  if Prefixe = 'ADL' then
    Suffixe := 'TIERSLIVRE'
  else Suffixe := 'TIERSFACTURE';

  Result :=TiersAuxiliaire(GetControlText('GP_' + Suffixe),false);
end;

procedure TOF_ComplPiece.SetLastError(Num: integer; ou: string);
begin
  if ou<>'' then SetFocusControl(ou);
  LastError:=Num;
  LastErrorMsg:=TexteMessage[LastError];
  TForm(Ecran).ModalResult:=0;
end;

procedure TOF_ComplPiece.CP_ADR_IDENTOnClick(Sender: TObject);
begin
  if TCheckBox(GetControl('CP_ADR_IDENT')).Checked then
  begin
    // Si c'est une adresse de livraison référencée dans les adresses
    if ExisteSQL('SELECT ADR_LIVR FROM ADRESSES WHERE ADR_REFCODE = "'+GetControlText('GP_TIERSLIVRE') +  '" AND ADR_NUMEROADRESSE = ' + GetControlText('GP_NUMADRESSELIVR') + ' AND ADR_LIVR="X"') then
    begin
      // Si c'est une adresse de livraison référencée mais pas une adresse de facturation
      if not ExisteSQL('SELECT ADR_FACT FROM ADRESSES WHERE ADR_REFCODE = "'+GetControlText('GP_TIERSLIVRE') + '" AND ADR_NUMEROADRESSE = ' + GetControlText('GP_NUMADRESSELIVR') + ' AND ADR_FACT="X"') then
      begin
        PGIBOX('Cette adresse n''est pas référencée comme une adresse de facturation');
        SetControlChecked('CP_ADR_IDENT',False);
        exit;
      end;
    end;
    SetControlEnabled('MNFAC', False);
    SetControlText('GP_NUMADRESSEFACT', GetControlText('GP_NUMADRESSELIVR'));
    SetControlText('GP_TIERSFACTURE', GetControlText('GP_TIERSLIVRE'));
    SetControlText('ADF_LIBTIERS', GetControlText('ADL_LIBTIERS'));
    SetControlText('ADF_LIBTIERS2', GetControlText('ADL_LIBTIERS2'));
    SetControlText('ADF_ADRESSE1', GetControlText('ADL_ADRESSE1'));
    SetControlText('ADF_ADRESSE2', GetControlText('ADL_ADRESSE2'));
    SetControlText('ADF_ADRESSE3', GetControlText('ADL_ADRESSE3'));
    SetControlText('ADF_CODEPOSTAL', GetControlText('ADL_CODEPOSTAL'));
    SetControlText('ADF_VILLE', GetControlText('ADL_VILLE'));
    SetControlText('ADF_PAYS', GetControlText('ADL_PAYS'));
    SetControlText('ADF_INCOTERM', GetControlText('ADL_INCOTERM'));
    SetControlText('ADF_MODEEXP', GetControlText('ADL_MODEEXP'));
    SetControlText('ADF_LIEUDISPO', GetControlText('ADL_LIEUDISPO'));
    SetControlText('ADF_REGION', GetControlText('ADL_REGION'));
    SetControlText('ADF_NUMEROCONTACT', GetControlText('ADL_NUMEROCONTACT'));
    SetControlText('ADF_CONTACT', GetControlText('ADL_CONTACT'));
    SetControlText('ADF_CONTACTPRENOM', GetControlText('ADL_CONTACTPRENOM'));
    SetControlText('ADF_EAN', GetControlText('ADL_EAN'));
    SetControlText('ADF_NIF', GetControlText('ADL_NIF'));

    SetAdresseFacturationEnabled(False);
  end
  else
  begin
    SetControlEnabled('mnFac', True);
    SetAdresseFacturationEnabled(True);

    if GetControlText('GP_NUMADRESSELIVR')='1' then
      SetControlText('GP_NUMADRESSEFACT', '2')
    else
      SetControlText('GP_NUMADRESSEFACT', '-2');
  end;
end;

procedure TOF_ComplPiece.SetAdresseFacturationEnabled(Saisissable: boolean);
Var
  ColorEnabled, ColorUnEnabled : Tcolor;
begin
  ColorEnabled := THEDIT(GetControl('GP_TIERSLIVRE')).Color;
  ColorUnEnabled := THEDIT(GetControl('GBFACTURE')).Color;

  SetControlEnabled('GP_TIERSFACTURE', Saisissable);
  SetControlEnabled('ADF_LIBTIERS', Saisissable);
  SetControlEnabled('ADF_LIBTIERS2', Saisissable);
  SetControlEnabled('ADF_ADRESSE1', Saisissable);
  SetControlEnabled('ADF_ADRESSE2', Saisissable);
  SetControlEnabled('ADF_ADRESSE3', Saisissable);
  SetControlEnabled('ADF_CODEPOSTAL', Saisissable);
  SetControlEnabled('ADF_VILLE', Saisissable);
  SetControlEnabled('ADF_PAYS', Saisissable);
  SetControlEnabled('ADF_INCOTERM', Saisissable);
  SetControlEnabled('ADF_MODEEXP', Saisissable);
  SetControlEnabled('ADF_LIEUDISPO', Saisissable);
  SetControlEnabled('ADF_REGION', Saisissable);
  SetControlEnabled('ADF_NUMEROCONTACT', Saisissable);
  SetControlEnabled('ADF_EAN', Saisissable);
  SetControlEnabled('ADF_NIF', Saisissable);

  if Saisissable then
  begin
    SetControlProperty('GP_TIERSFACTURE'  , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_LIBTIERS'     , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_LIBTIERS2'    , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_ADRESSE1'     , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_ADRESSE2'     , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_ADRESSE3'     , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_CODEPOSTAL'   , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_VILLE'        , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_PAYS'         , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_INCOTERM'     , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_MODEEXP'      , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_LIEUDISPO'    , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_REGION'       , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_NUMEROCONTACT', 'COLOR', ColorEnabled);
    SetControlProperty('ADF_EAN'          , 'COLOR', ColorEnabled);
    SetControlProperty('ADF_NIF'          , 'COLOR', ColorEnabled);
  end
  else
  begin
    SetControlProperty('GP_TIERSFACTURE'  ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_LIBTIERS'     ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_LIBTIERS2'    ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_ADRESSE1'     ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_ADRESSE2'     ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_ADRESSE3'     ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_CODEPOSTAL'   ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_VILLE'        ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_PAYS'         ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_INCOTERM'     ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_MODEEXP'      ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_LIEUDISPO'    ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_REGION'       ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_NUMEROCONTACT','COLOR', ColorUnEnabled);
    SetControlProperty('ADF_EAN'          ,'COLOR', ColorUnEnabled);
    SetControlProperty('ADF_NIF'          ,'COLOR', ColorUnEnabled);
  end;
end;

Initialization
registerclasses([TOF_ComplLigne, TOF_ComplPiece, TOF_SaisCond]) ;
InitTOFPiece();
end.
