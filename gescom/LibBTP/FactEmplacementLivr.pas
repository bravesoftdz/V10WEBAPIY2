unit FactEmplacementLivr;

interface
uses sysutils,classes,windows,messages,controls,forms,hmsgbox,stdCtrls,clipbrd,
     HCtrls,SaisUtil,HEnt1,Ent1,EntGC,UtilPGI,UTOB,HTB97,Menus,
     ParamSoc,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  fe_main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  FactVariante,
{$IFDEF BTP}
     BTPUtil,CalcOlegenericBTP,
{$ENDIF}
     HRichOLE,uEntCommun;

const MAXITEMS = 5;
type

TmodeLivraison = (tmlchantier,tmldepot,tmlfournisseur);

TDestinationLivraison = class
  private
    fCreatedPop : boolean;
    FF : TForm;
    fusable : boolean;
    fPositionnable : boolean;
    fgrid : THGrid;
    POPGS : TPopupMenu;
    TOBPiece,TOBLigPiece : TOB;
    TOBAffaire,TOBTiers : TOB;
    TOBAdresses,TOBContact : TOB;
    TOBL : TOB;
    fMode : TModeLivraison;
    MesMenuItem: array[0..MAXITEMS] of TMenuItem;
    fMaxItems : integer;
    fNaturePiece : String;
    TheRegoupMenu : TMenuItem;

    //Modif FV
    fIdentifiantWol : integer; // -1 livraison sur chantier 0 autre
    SaveIdWol       : Integer;

    procedure SetPositionnable(const Value: boolean);
		procedure GSLivraisonChantier (Sender : TObject);
    procedure GSLivraisonDepot (Sender : TObject);
    procedure GSDispoFournisseur (Sender : TObject);
    procedure DefiniMenuPop(Parent: Tform);
    procedure Setdocument(const Value: TOB);
    procedure SetAffaire(const Value: TOB);
    function  IsDocumentOK: boolean;
    function  DocumentUsitable: boolean;
    function  IsProvenanceVente(Reference: string): boolean;
    procedure ActiveIt (active : boolean);
    procedure SetUsable(const Value: boolean);
  	property  Usable : boolean read fusable write SetUsable;
  	property  Positionnable : boolean read fPositionnable write SetPositionnable;
    procedure DetermineTypeLivr;
    procedure SetGrille(const Value: THgrid);
    procedure SetLigneReferenceLiv(TOBL: TOB);
    procedure SetAdresseLivChantier(MemoTemp: TStringList);
    procedure SetAdresseDepot(TOBL: TOB; memoTemp: TStringList);
    procedure DefiniReferenceLivraison;
    procedure SetAdresseLivAffaire;
    procedure PositionneLivLignes;
    function  IsPositionnable: boolean;
    function  Naturepieceusable(NaturePiece: string): boolean;
    procedure ChangementAdresse(TOBL: TOB);

  public
    constructor create (TT : TForm);
    destructor  destroy ; override;
    property    document : TOB read TobPiece write Setdocument;
    property    Tiers : TOB read TobTiers write TobTiers;
    property    affaire :TOb read TOBaffaire write Setaffaire;
    property    grid : THgrid read fgrid write SetGrille;
    procedure   ChangeAffaire ;
    procedure   ChangeTiers ;
    procedure   ChangeDepot;
    procedure   SetDestLivrDefaut(TOBL: TOB);
    procedure InitModeLivraison;
    procedure SetAdresseLivDepot(TOBL: TOB);
end;

function IsLigneReferenceLivr (TOBL : TOB) : Boolean;

implementation
uses facture,FactGrpBtp,FactUtil,nomenUtil,factcomm,
		 FactTOB, FactPiece, FactArticle,FactAdresseBTP;

function IsLigneReferenceLivr (TOBL : TOB) : Boolean;
begin
	result := (TOBL.GetValue('GL_TYPELIGNE') = 'RL');
end;
{ TDestinationLivraison }

constructor TDestinationLivraison.create(TT: TForm);
var ThePop : Tcomponent;
begin
  FF := TT;
  fusable := false;
  fPositionnable := false;
  if FF is TFFacture then ThePop := TFFacture(TT).Findcomponent  ('POPBTP')
                     else ThePop := nil;
  if ThePop = nil then
  BEGIN
    // pas de menu BTP trouve ..on le cree
    POPGS := TPopupMenu.Create(TT);
    POPGS.Name := 'POPBTP';
    fCreatedPop := true;
  END else
  BEGIN
    fCreatedPop := false;
    POPGS := TPopupMenu(thePop);
  END;
  DefiniMenuPop(TT);
  TOBAdresses := TOB.Create ('LES ADRESSES', nil,-1);
  TOBCOntact := TOB.Create ('CONTACT',nil,-1);
  // initialisation par defaut
  InitModeLivraison;
end;

procedure TDestinationLivraison.InitModeLivraison;
var LivBesoinParFour : boolean;
begin
  // par défaut
   fMode := tmldepot;
   fIdentifiantWol := 0;
   LivBesoinParFour  := not GetParamSocSecur('SO_BTLIVBESOINDEF',false);
  if GetParamSoc ('SO_BTLIVCHANTIER') then
  begin
    if ((TFFacture(FF).NewNature<>'CBT') or
       ((TFFacture(FF).NewNature='CBT') and (LivBesoinParFour))) then
    begin
      fMode := tmlchantier;
      fIdentifiantWol := -1;
    end;
  end;

  //Modif FV
  SaveIdWol := fIdentifiantWol;

end;

procedure TDestinationLivraison.DefiniMenuPop(Parent: Tform);
var Indice,ItemMenu : integer;
begin
  fmaxitems := 0; Indice := 0;
  if not fcreatedPop then
  begin
    MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
    with MesMenuItem[fmaxitems] do
    begin
      Caption := '-';
      visible := false;
    end;
		POPGS.Items.Add (MesMenuItem[fmaxitems]);
    inc (fmaxitems);
  end;

  // CTRL+ALT+C = livraison a l'adresse de l'affaire
  MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
  with MesMenuItem[fmaxitems] do
    begin
    Caption := TraduireMemoire ('Sur le chantier');  // par défaut
    Name := 'BLIVRCHANTIER';
    OnClick := GSLivraisonChantier;
    visible := false;
    end;
  MesMenuItem[fmaxitems].ShortCut := ShortCut( Word('C'), [ssCtrl,ssAlt]);
  Indice := fMaxItems;
  inc (fmaxitems);

  // CTRL+ALT+M = livraison au dépot
  MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
  with MesMenuItem[fmaxitems] do
    begin
    Caption := TraduireMemoire ('Au dépot');  // par défaut
    Name := 'BLIVRDEPOT';
    OnClick := GSLivraisonDepot;
    visible := false;
    end;
  MesMenuItem[fmaxitems].ShortCut := ShortCut( Word('D'), [ssCtrl,ssAlt]);
  inc (fmaxitems);

  // CTRL+ALT+D = Disposition chez vous
  MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
  with MesMenuItem[fmaxitems] do
    begin
    Caption := TraduireMemoire ('A disposition chez fournisseur');  // par défaut
    Name := 'BDISPOFOU';
    OnClick := GSDispoFournisseur;
    visible := false;
    end;
  MesMenuItem[fmaxitems].ShortCut := ShortCut( Word('F'), [ssCtrl,ssAlt]);
  inc (fmaxitems);

  // Sousmenu
  MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
  with MesMenuItem[fmaxitems] do
    begin
    Caption := TraduireMemoire ('Livraison fournisseur (Document)');
    Name := 'BLIVRAISON';
    visible := false;
    end;
  TheRegoupMenu :=MesMenuItem[fmaxitems];
  inc (fmaxitems);

  for ItemMenu := Indice to Indice+ 2 do
  begin
  	TheRegoupMenu.Add (MesMenuItem[ItemMenu]);
  end;

	POPGS.Items.Add (TheRegoupMenu);

end;

destructor TDestinationLivraison.destroy;
var indice : integer;
begin
  for Indice := 0 to fmaxitems -1 do
  begin
    MesMenuItem[Indice].Free;
  end;
  if fcreatedPop then POPGS.free;
  TOBAdresses.free;
  TOBCOntact.free;
  inherited;
end;

function TDestinationLivraison.IsProvenanceVente (Reference : string) : boolean;
var cledoc : R_CLEDOC;
begin
	result := false;
	DecodeRefPiece (Reference,Cledoc);
  if GetInfoParPiece(Cledoc.NaturePiece, 'GPP_VENTEACHAT') = 'VEN' Then result := true;
end;

function TDestinationLivraison.DocumentUsitable : boolean;
var Indice : integer;
		TOBL : TOB;
begin

  //	result := true;
  result := Naturepieceusable (TOBPiece.getValue('GP_NATUREPIECEG')); if result = false then exit;

	for indice := 0 to TOBPiece.detail.count -1 do
  begin
  	TOBL := TOBpiece.detail[Indice];
    if TOBL.GetValue ('GL_TYPELIGNE') <> 'ART' Then continue;
    if TOBL.GetValue ('GL_PIECEORIGINE')='' then continue;
    //if IsProvenanceVente (TOBL.GetValue ('GL_PIECEORIGINE')) then BEGIN result := false; break; END;
    if TOBL.GetValue ('GL_QTEFACT') <> TOBL.GetValue ('GL_QTERESTE') then result := false;
    // --- GUINIER ---
    if CtrlOkReliquat(TOBL, 'GL') then
    begin
      if TOBL.GetValue ('GL_MONTANTHTDEV') <> TOBL.GetValue ('GL_MTRESTE') then result := false;
    end;
    if not Result then  Break;
  end;

end;

function TDestinationLivraison.ISDocumentOK : boolean;
begin
	result := true;
  if TOBPiece = nil then BEGIN Result := false; Exit ;END;
//  if not ((TFFacture(FF).Action = taCreat) or (TFFacture(FF).Action = TaModif)) then BEGIN result := false; exit; END;
  if TFFacture(FF).Action = TaConsult then BEGIN result := false; exit; END;
(*
  if ((TFFacture(FF).TransfoPiece = True) and (TFFacture(FF).Action = TaModif)) then BEGIN result := false; exit; END;
*)
  if TOBPiece.GetValue('GP_AFFAIRE')='' then BEGIN Result := false; Exit ;END;
//  if not DocumentUsitable then BEGIN result := false;Exit; END;
//	if Pos(TOBPiece.getValue('GP_NATUREPIECEG'),GetPieceAchat (false,true,false)) = 0 then BEGIN result := false; Exit; END;
end;

function TDestinationLivraison.Naturepieceusable (NaturePiece : string) : boolean;
begin
	result := true;
	if (Pos(NaturePiece,GetPieceAchat (false,true,false)) = 0) and
  	 (Pos(NaturePiece,'CBT;') = 0) then result := false;
end;

function TDestinationLivraison.IsPositionnable : boolean;
begin
	result := True;
  if TOBPiece = nil then BEGIN Result := false; Exit ;END;
  if TFFacture(FF).Action = TaConsult then BEGIN result := false; exit; END;
(*
  if ((TFFacture(FF).TransfoPiece = True) and (TFFacture(FF).Action = TaModif)) then BEGIN result := false; exit; END;
*)
  if TOBPiece.GetValue('GP_AFFAIRE')='' then BEGIN Result := false; Exit ;END;
  if not DocumentUsitable then BEGIN result := false;Exit; END;
  if not Naturepieceusable (TobPiece.getValue('GP_NATUREPIECEG')) then begin result := false; exit; end;
end;

procedure TDestinationLivraison.GSLivraisonChantier(Sender: TObject);
begin
	fmode := tmlchantier;
  fIdentifiantWol := -1;
  DefiniReferenceLivraison;
end;

procedure TDestinationLivraison.GSLivraisonDepot(Sender: TObject);
begin
	fmode := tmldepot;
  fIdentifiantWol := 0;
  DefiniReferenceLivraison;
end;



procedure TDestinationLivraison.SetAdresseLivDepot(TOBL : TOB);
begin
	fmode := tmldepot;
  fIdentifiantWol := 0;
  DefiniReferenceLivraison;
end;

procedure TDestinationLivraison.DefiniReferenceLivraison;
var Indice : integer;
		TOBL : TOB;
    ARow , Acol : integer;
begin

	if TOBPiece.GetValue('GP_TIERS')='' then exit;

  if (Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'PBT;CBT;DE;PRO;CC;BLC;FAC;LBT') > 0) then
     BEGIN
     TOBPiece.putValue('GP_IDENTIFIANTWOT', fIdentifiantWol);
     exit;
     END;

  if (Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'CF') > 0) then
  BEGIN
    TOBPiece.putValue('GP_IDENTIFIANTWOT', fIdentifiantWol);
  END;

  TOBL := TOBPiece.findFirst(['GL_TYPELIGNE'],['RL'],true);

  if TOBL = nil then
     begin
  	 if TOBPiece.detail.count > 0 then
        TFFActure(FF).ClickInsert(1)
     else
        CreerTOBLignes (fgrid,TOBpiece,TOBTiers,TOBAffaire,1);
     TOBL := GetTOBLigne(TOBpiece,1);
     if TOBL = nil then exit;
     PieceVersLigne(TOBPiece, TOBL);
     TOBL.PutValue('GL_NUMLIGNE', 1);
     TOBL.PutValue('RECALCULTARIF', '-');
     TOBL.PutValue('GL_PRIXPOURQTE', 1);
     TOBL.PutValue('GL_CODECOND', '');
     TOBL.PutValue('GL_INDICENOMEN', 0);
     TOBL.PutValue('GL_TYPELIGNE', 'RL');
     TOBL.PutValue('GL_TYPEDIM', 'NOR');
     TOBL.PutValue('GL_DATELIVRAISON', iDate1900);
     InitialiseComm(TOBL, False);
     TOBL.PutValue('GL_NUMORDRE', 0) ;
     TOBL.PutValue('GL_DEPOT', TOBPiece.GetValue('GP_DEPOT')) ;
     end;

  Indice := TOBL.getValue('GL_NUMLIGNE');

  SetLigneReferenceLiv (TOBL);

  TFFacture(FF).AfficheLaLigne (Indice);

(*
  if (Fgrid.row = 0) or (FGrid.col = 0) then exit;
  ARow := 1;
  Acol := fgrid.Col;
  TFFacture(FF).GoToLigne (Arow,Acol);
*)
end;

procedure TDestinationLivraison.DetermineTypeLivr;
var TOBL : TOB;
begin
(*
	fMode := tmlDepot;
	fIdentifiantWol := 0;
*)
  TOBL := TOBPiece.findFirst(['GL_TYPELIGNE'],['RL'],true);
  if TOBL = nil then Exit;
  if TOBL <> nil then fIdentifiantWol := TOBL.GetValue('GL_IDENTIFIANTWOL');

  //Modif FV
  SaveIdWol := fIdentifiantWol;

  if fidentifiantWol = -1 then
     begin
  	 fmode := tmlchantier;
     end
  else
     begin
  	 if Pos(TraduireMemoire('Dépot'),TOBL.GetValue('GL_BLOCNOTE')) > 0 then
        fmode := tmldepot
     else
        fmode := tmlfournisseur;
     end;

end;

procedure TDestinationLivraison.SetDestLivrDefaut (TOBL : TOB);
begin
//	if (not fusable) and (not fPositionnable) then exit;
	if TOBL = nil then exit;

	TOBL.PutValue('GL_IDENTIFIANTWOL',fIdentifiantWol);
end;

procedure TDestinationLivraison.SetAffaire(const Value: TOB);
begin

  TobAffaire := value;

  if tobAffaire.GetString('AFF_ETATAFFAIRE')='TER' then
    SetPositionnable (false)
  else
    SetAdresseLivAffaire;

end;

procedure TDestinationLivraison.SetAdresseLivAffaire;
begin
  TOBAdresses.clearDetail;
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Livraison}
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Facturation}
  end else
  begin
    TOB.Create('ADRESSES', TOBAdresses, -1); {Livraison}
    TOB.Create('ADRESSES', TOBAdresses, -1); {Facturation}
  end;
	LivAffaireVersAdresses (TOBAffaire,TOBAdresses,TOBPiece,false);
end;

procedure TDestinationLivraison.Setdocument(const Value: TOB);
begin
  TOBPiece := Value;
  if TOBPiece = nil then BEGIN Usable := False; Exit; END;
  fNaturePiece := TOBPiece.getValue('GP_NATUREPIECEG');
  {Controle si le document n'est pas partielement livré ou si la nature de piece est autorisée }
  Usable := ISDocumentOK;
  Positionnable := IsPositionnable;

  if (fMode=tmlchantier) and (Usable) then
     fIdentifiantWol := -1
  else
     fIdentifiantWol := 0;

  //Modif FV
  SaveIdWol := fIdentifiantWol;

  if fusable then DetermineTypeLivr;

end;

procedure TDestinationLivraison.ActiveIt (active : boolean);
var Indice : integer;
begin
  for Indice := 0 to fmaxitems -1 do
  begin
  	if MesMenuItem [Indice] <> nil then
    begin
    	MesMenuItem[Indice].visible := active;
    end;
  end;
end;

procedure TDestinationLivraison.SetUsable(const Value: boolean);
begin
  fusable := Value;
  if fusable then ActiveIt(True) else ActiveIt (False);
end;

procedure TDestinationLivraison.ChangeAffaire;
begin
	SetAdresseLivAffaire;
//  if not fusable then exit;
  Usable := ISDocumentOK;
  Positionnable := IsPositionnable;
  
  if (fMode=tmlchantier) and (Usable) then
     fIdentifiantWol := -1
  else
     fIdentifiantWol := 0;

  if (Usable) and (fnaturepiece<> '') and (Pos(fNaturePiece,'ETU;FRC;DBT;FBT;ABT;ABP;CBT;LBT;')=0) then
  begin
  	DefiniReferenceLivraison;
  end;
end;

Procedure TDestinationLivraison.ChangeDepot;
var TOBL : TOB;
begin

  TOBL := TOBPiece.findFirst(['GL_TYPELIGNE'],['RL'],true);

  If TOBL = Nil then exit;

  fIdentifiantWol := TOBL.GetValue('GL_IDENTIFIANTWOL');

  if      fIdentifiantWol  = -1 then fmode := tmlchantier
  else if fIdentifiantWol  = 0  then fmode := tmldepot
  else fmode := tmlfournisseur;

  ChangementAdresse(TOBL);

end;

procedure TDestinationLivraison.PositionneLivLignes;
var indice : integer;
		TOBL : TOB;
begin
	for Indice := 0 to TOBPiece.detail.count -1 do
  begin
  	TOBL := TOBPiece.detail[indice];
    TOBL.PutValue('GL_IDENTIFIANTWOL',fIdentifiantWol);
  end;
end;

procedure TDestinationLivraison.GSDispoFournisseur(Sender: TObject);
begin
	fmode := tmlfournisseur;
  fIdentifiantWol := 2;
  if (fnaturepiece<> '') and (Pos(fNaturePiece,'ETU;FRC;DBT;FBT;ABT;ABP;CBT;LBT;')=0) then
  begin
  	DefiniReferenceLivraison;
  end;
end;

procedure TDestinationLivraison.SetGrille(const Value: THgrid);
begin
  fgrid := Value;
end;

Procedure TDestinationLivraison.ChangementAdresse(TOBL : TOB);
var Chaine : string;
    MemoTemp : TStringList;
begin

  MemoTemp := TStringList.create;
  MemoTemp.text := '';

  TOBL.putvalue('GL_AFFAIRE',TOBPiece.GetValue('GP_AFFAIRE'));

  if copy(TOBPiece.GetValue('GP_AFFAIRE'),1,1) = 'P' then
	   Chaine := Traduirememoire('Appel d''Offre')
  Else if copy(TOBPiece.GetValue('GP_AFFAIRE'),1,1) = 'I' then
	   Chaine := Traduirememoire('Contrat')
  Else if copy(TOBPiece.GetValue('GP_AFFAIRE'),1,1) = 'W' then
	   Chaine := Traduirememoire('Appel')
	Else
     Chaine := Traduirememoire('Chantier');

  {$IFDEF BTP}
  Chaine := Chaine +' '+ BTPCodeAffaireAffiche (TOBPiece.GetValue('GP_AFFAIRE'));
  {$ELSE}
  Chaine := Chaine +' '+ CodeAffaireAffiche (TOBPiece.GetValue('GP_AFFAIRE'));
  {$ENDIF}
  TOBL.putvalue('GL_LIBELLE',Chaine);

  if fMode = tmlChantier then
  begin
    SetAdresseLivChantier (MemoTemp);
  end else if fMode = tmlDepot then
  begin
    SetAdresseDepot (TOBL,MemoTemp);
  end else
  begin
    MemoTemp.add (TraduireMemoire('Mise à disposition dans vos locaux'));
  end;

  TOBL.putValue('GL_BLOCNOTE',MemoTemp.Text);

  MemoTemp.free;

end;

procedure TDestinationLivraison.SetLigneReferenceLiv (TOBL : TOB);
begin

  ChangementAdresse(TOBL);

  //Modif FV
  //Si on modifie la livraison de chantier alors que celle ci était sur le chantier
  //on pose la question de savoir si oui ou non on veut la conserver !!!
  //if (SaveIdWol = -1) And (fIdentifiantWol = 0) then
  if (SaveIdWol <> fIdentifiantWol) and (fIdentifiantWol <> -1)  then
  Begin
(*
    if (PGIAsk(TraduireMemoire('Voulez-vous conserver la mise à jour automatique des consommations ?'), 'Modification Lieu de Livraison') = mryes) then
    begin
      fIdentifiantWol := -1;
      TOBL.PutValue('GL_IDENTIFIANTWOL',fIdentifiantWol);
    end
    else
    begin
      PositionneLivLignes;
    end;
*)
    if (PGIAsk(TraduireMemoire('Voulez-vous conserver la mise à jour automatique des consommations ?'), 'Modification Lieu de Livraison') = mryes) then
    begin
      fIdentifiantWol := -1;
    end;
    PositionneLivLignes;
  end
  else
  begin
    PositionneLivLignes;
  end;

  SaveIdWol := fIdentifiantWol;

end;

procedure TDestinationLivraison.SetAdresseLivChantier (MemoTemp : TStringList);
var chaine : string;
begin

	MemoTemp.add (TraduireMemoire('Livraison à l''adresse suivante :'));

  if GetParamSoc('SO_GCPIECEADRESSE') then
     Begin
     if TobAdresses.detail[0].getvalue('GPA_LIBELLE') <> '' then
        MemoTemp.add(TobAdresses.detail[0].getvalue('GPA_LIBELLE'));
     if TobAdresses.detail[0].getvalue('GPA_LIBELLE2') <> '' then
        MemoTemp.add(TobAdresses.detail[0].getvalue('GPA_LIBELLE2'));
     if TobAdresses.detail[0].getvalue('GPA_ADRESSE1') <> '' then
        MemoTemp.add(TobAdresses.detail[0].getvalue('GPA_ADRESSE1'));
     if TobAdresses.detail[0].getvalue('GPA_ADRESSE2') <> '' then
        MemoTemp.add(TobAdresses.detail[0].getvalue('GPA_ADRESSE2'));
     if TobAdresses.detail[0].getvalue('GPA_ADRESSE3') <> '' then
        MemoTemp.add(TobAdresses.detail[0].getvalue('GPA_ADRESSE3'));
     if TobAdresses.detail[0].getvalue('GPA_CODEPOSTAL') <> '' then
        MemoTemp.add(TobAdresses.detail[0].getvalue('GPA_CODEPOSTAL') + ' '+ TobAdresses.detail[0].getvalue('GPA_VILLE'));
     if TobAdresses.detail[0].getvalue('GPA_NUMEROCONTACT') <> 0 then
        begin
        if getContact (TOBContact,TOBAffaire.getValue('AFF_TIERS'),TobAdresses.detail[0].getvalue('GPA_NUMEROCONTACT')) then
           begin
           chaine := 'Contact : '+rechDom('TTCIVILITE',TOBCOntact.getValue('C_CIVILITE'),false)  + ' '+
                      TOBContact.GetValue('C_PRENOM')+' '+ TOBContact.GetValue('C_NOM');
           if TOBContact.GetValue('C_TELEPHONE') <> '' then
              begin
              chaine := Chaine + ' '+Traduirememoire('Tel :')+' '+ TOBContact.GetValue('C_TELEPHONE');
              end;
              MemoTemp.add (chaine);
           end;
        end;
  End ELSE
     BEGIN
     if TobAdresses.detail[0].getvalue('ADR_libelle') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_libelle'));
     if TobAdresses.detail[0].getvalue('ADR_libelle2') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_libelle2'));
     if TobAdresses.detail[0].getvalue('ADR_ADRESSE1') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_ADRESSE1'));
     if TobAdresses.detail[0].getvalue('ADR_ADRESSE2') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_ADRESSE2'));
     if TobAdresses.detail[0].getvalue('ADR_ADRESSE3') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_ADRESSE3'));
     if TobAdresses.detail[0].getvalue('ADR_CODEPOSTAL') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_CODEPOSTAL') +' '+ TobAdresses.detail[0].getvalue('ADR_VILLE'));
     if TobAdresses.detail[0].getvalue('ADR_NUMEROCONTACT') <> 0 then
        begin
        if getContact (TOBContact,TOBAffaire.getValue('AFF_TIERS'),TobAdresses.detail[0].getvalue('ADR_NUMEROCONTACT')) then
           begin
           chaine := 'Contact : '+rechDom('TTCIVILITE',TOBCOntact.getValue('C_CIVILITE'),false)  + ' '+
                     TOBContact.GetValue('C_PRENOM')+' '+ TOBContact.GetValue('C_NOM');
           if TOBContact.GetValue('C_TELEPHONE') <> '' then
              begin
              chaine := Chaine + ' '+Traduirememoire('Tel :')+' '+ TOBContact.GetValue('C_TELEPHONE');
              end;
           MemoTemp.add (chaine);
           end;
        end;
     END;

end;


procedure TDestinationLivraison.SetAdresseDepot (TOBL : TOB;memoTemp : TStringList);
begin
	getAdresseSoc (TOBL,MemoTemp);
end;

procedure TDestinationLivraison.ChangeTiers;
begin
  Usable := ISDocumentOK;
  Positionnable := IsPositionnable;
	if (fMode=tmlchantier) and (Usable) then fIdentifiantWol := -1 else fIdentifiantWol := 0;
  if (Usable) and (fnaturepiece<> '') and (Pos(fNaturePiece,'ETU;FRC;DBT;FBT;ABT;ABP;CBT;LBT;')=0) then
  begin
  	DefiniReferenceLivraison;
  end;
end;

procedure TDestinationLivraison.SetPositionnable(const Value: boolean);
begin
	TheRegoupMenu.enabled := value;
  fPositionnable := Value;
end;

end.
