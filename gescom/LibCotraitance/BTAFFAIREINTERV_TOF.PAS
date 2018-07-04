{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/03/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTAFFAIREINTERV ()
Mots clefs ... : TOF;BTAFFAIREINTERV
*****************************************************************}
Unit BTAFFAIREINTERV_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul, MainEagl,
{$ENDIF}
     uTob,
     Vierge,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     HPanel,
     graphics,
     ImgList,
     ExtCtrls,
     grids,
     types,
     TiersUtil,
     UTOF ;

Type
  TOF_BTAFFAIREINTERV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	GS                : ThGRid;
    Binsert           : TToolbarbutton97;
    Bdelete           : TToolbarButton97;
    Bok               : TToolbarButton97;
    Bvalider          : TToolbarbutton97;
    BaNNUL            : TToolbarbutton97;
    BFerme            : TToolbarbutton97;
    PSaisie           : THpanel;
    fListeSaisie      : string;
    fLibTiers         : String;
    fTypeInterv       : string;
  	fModegestion      : string;   // COTRAITANCE ou SOUSTRAITANCE
    fRoleEntreprise   : string;   // MANDATAIRE ou COTRAITANTE
    fAffaire          : string;   // Code de l'affaire
    fTOBAffaireInterv : TOB;      // TOb des intervenants
    fLastRow          : integer;
    fMaxInterv        : integer;
    BAI_NUMEROCONTACT : Thedit;
    BAI_TIERSCLI      : Thedit;
    BAI_TIERSFOU      : Thedit;
    BAI_NUMERORIB     : THEdit;
    BAI_TYPEINTERV    : THValComboBox;

    ImTypeArticle     : TImageList;
    TOBL              : TOB;
    FExiting          : boolean;
    //
    FColNames         : string;
    FNomTable         : string;
    FLien             : string;
    FSortBy           : string;
    FLargeur          : string;
    FAlignement       : string;
    FParams           : string;
    FPerso            : string;
    ErrorMsg          : String;
    NatureAuxi        : String;
    FTitre            : Hstring;
    Title             : hString;
    NC                : Hstring;
    OkTri             : boolean;
    OkNumCol          : Boolean;
    nbcolsinliste     : integer;
    //
    fInserting        : boolean;

    procedure LocalCreate ;
    procedure LocalDestroy;
    procedure DecodeArgs(S : string);
    procedure ChargeLesIntervenants;
    procedure ChargeLagrille;
    procedure SetEvents(active : boolean);
    procedure Positionne(Arow : integer);
    procedure GSRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSDblclick (Sender : Tobject);
    procedure GSPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);

  	procedure ResetInfo;
  	procedure SetInfoligne (TOBL : TOB);
    procedure BdeleteClick(Sender : Tobject);
    procedure BinsertClick (Sender : Tobject);
    procedure SetEventsSaisie (Etat : boolean);
    procedure BOkClick (Sender : Tobject);
    procedure BAnnulClick (Sender : TObject);
    procedure ChangeTzTiersSaisie;
  	procedure TIERSCLIEnter(Sender : Tobject);
  	procedure TIERSCLIEXit (Sender : Tobject);
  	procedure TIERSFOUEnter(Sender : Tobject);
  	procedure TIERSFOUEXit(Sender : Tobject);
    procedure TIERSCLIElipsisClick (Sender : TObject);
    procedure TIERSFOUElipsisClick (Sender : TObject);
    procedure ContactClick (Sender : TObject);
    procedure ContactExit (Sender : tobject);
    procedure CodeBQExit(Sender: TObject);
    procedure CodeBQOnChange(Sender: TObject);
    procedure SetparamGrille;
    procedure DefiniGrille;
    procedure AffichelaGrille(Entete : boolean=false);
    procedure AfficheLaLigne(Indice: integer; Entete: boolean=false);
    procedure DefiniEcran;
    procedure ChargeListImage;
    Procedure RIBElipsisClick(Sender: TObject);
    procedure Charge_InfoBQE(TOBBQE: TOB; indice : Integer);
    procedure SetInfoBqe(TOBL: TOB);

    function ControleExist (TOBL : TOB ;Fournisseur : string; TypeIntervenant : string) : integer;
    function  GetTOBLigne(Arow: integer): TOB;
    function InsertTOBligne : TOB;
    function OnBeforeUpdate(var TypeErreur: string): Boolean;
    function RecupTypeGraph(TOBL: TOB): integer;

  end ;

Implementation
uses FactTiers,BTPUtil;
procedure TOF_BTAFFAIREINTERV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAIREINTERV.OnDelete ;
begin
  Inherited ;
end ;

function TOF_BTAFFAIREINTERV.OnBeforeUpdate(var TypeErreur : string) : Boolean ;

begin

  Result := True;

  if (BAI_NUMEROCONTACT.Text = '') OR (BAI_NUMEROCONTACT.Text = '0') then
  begin
    BAI_NUMEROCONTACT.SetFocus;
    ErrorMsg := 'Attention le contact n''est pas renseigné.' + CHR(10) + 'Confirmez-vous la validation';
    TypeErreur := 'N';
    Result := False;
  end;   

end;

procedure TOF_BTAFFAIREINTERV.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAIREINTERV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAIREINTERV.OnArgument (S : String ) ;
begin
  Inherited ;
  LocalCreate;
  DecodeArgs(S);
  definiEcran;
  ChargeLesIntervenants;
  fLastRow := 1;
  ChargeLagrille;
  SetEvents(true);
  SetEventsSaisie (False);
  Positionne(flastRow);
end ;

procedure TOF_BTAFFAIREINTERV.OnClose ;
begin
	LocalDestroy;
  Inherited ;
end ;

procedure TOF_BTAFFAIREINTERV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAIREINTERV.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAIREINTERV.LocalCreate;
begin

  fTOBAffaireInterv := TOB.create ('LES INTERV AFF',nil,-1);

  ImTypeArticle := TImageList.create(ecran);
  ImTypeArticle.ImageType := itImage;

	fModegestion    := '';
  fRoleEntreprise :='';
  fAffaire        := '';
  fInserting      := false;

  GS        := ThGrid(GetControl('GS'));
  PSaisie   := THPanel (getControl('PSAISIE'));

  Binsert   := TToolbarButton97 (GetControl('Binsert'));
  Bdelete   := TToolbarButton97 (GetControl('BDelete'));
  Bok       := TToolbarButton97 (GetControl('BOk'));
  Bvalider  := TToolbarbutton97 (getControl('BValider'));
  BANNUL    := TToolbarbutton97 (getControl('BANNULE'));
  BFerme    := TToolbarbutton97 (getControl('BFerme'));
  //
  BAI_TIERSCLI      := Thedit (getControl('BAI_TIERSCLI'));
  BAI_TIERSFOU      := Thedit(getControl('BAI_TIERSFOU'));
  BAI_NUMEROCONTACT := Thedit(getControl('BAI_NUMEROCONTACT'));
  BAI_NUMERORIB     := THEdit(GetControl('BAI_NUMERORIB'));
  //
  BAI_TYPEINTERV    := THvalComboBox(getControl('BAI_TYPEINTERV'));
  //

  BAI_TIERSCLI.OnEnter              := TIERSCLIEnter;
  BAI_TIERSCLI.OnExit               := TIERSCLIEXit;
  BAI_TIERSCLI.OnElipsisClick       := TIERSCLIElipsisClick;
  //
  BAI_TIERSFOU.OnEnter              := TIERSFOUEnter;
  BAI_TIERSFOU.OnExit               := TIERSFOUEXit;
  BAI_TIERSFOU.OnElipsisClick       := TIERSFOUElipsisClick;
  //
  BAI_NUMEROCONTACT.OnElipsisClick  := ContactClick;
  BAI_NUMEROCONTACT.OnExit          := ContactExit;
  //
  BAI_NUMERORIB.OnElipsisClick      := RIBElipsisClick;
  BAI_NUMERORIB.OnChange            := CodeBQOnChange;
  BAI_NUMERORIB.OnExit              := CodeBQExit;
  //
  fListeSaisie := 'BTAFFAIREINTERV';
  SetParamgrille;
  DefiniGrille;

end;

procedure TOF_BTAFFAIREINTERV.DefiniGrille;
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep,Okimg : boolean;
    dec : integer;
    indice : integer;
    FFQTE,FFLQTE : string;
begin
	FFQTE := '###';
  if V_PGI.OkDecQ > 0 then
  begin
    FFQTE := FFQTE+'0.';
    for indice := 1 to V_PGI.OkDecQ do
    begin
      FFQTE := FFQTE + '0';
    end;
  end;

  GS.ColCount := NbColsInListe;
  st := fColNames;
  lesalignements := Falignement;
  lestitres := FTitre;
  leslargeurs := flargeur;

  for indice := 0 to nbcolsInListe -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    alignement := ReadTokenSt(lesalignements);
    lalargeur := readtokenst(leslargeurs);
    letitre := readtokenst(lestitres);
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
    GS.cells[Indice,0] := leTitre;
    GS.ColNames [Indice] := Nam;
    if copy(Alignement,1,1)='G' then GS.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then GS.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then GS.ColAligns[indice] := taCenter;
    okImg := (copy(Alignement,8,1)='X');
    if OkVisu then
    begin
  		GS.ColWidths[indice] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
    end else
    begin
    	GS.ColWidths[indice] := -1;
    end;
    
    if Nam='BAI_TYPEINTERV' then
    begin
    	OkImg := true;
    end;

    if (OkLib) or (okImg) then
    begin
    	GS.ColFormats[indice] := 'CB=' + Get_Join(Nam);
      if OkImg then
      begin
      	GS.ColDrawingModes[Indice]:= 'IMAGE';
      end;
    end
    else if (Dec<>0) or (Sep) then
    begin
    	if OkNulle then
      begin
				FFLQte := FFQTE+';-'+FFQTE+'; ;';
      	GS.ColFormats[indice] := FFLQTE ;
      end else
      begin
      	GS.ColFormats[indice] := FF ;
      end;
    end;
  end ;
end;

procedure TOF_BTAFFAIREINTERV.SetparamGrille;
var lelement,laliste : string;
begin
  // récupération du paramétrage général des grilles

  ChargeHListe (fListeSaisie, FNomTable, FLien, FSortBy, fColNames, FTitre,
                FLargeur, FAlignement, FParams, title, NC, FPerso, OkTri, OkNumCol);
  //
  // Ajout de la colonne désignation
  fColNames := fColNames+'TLIBTIERS;';
  Falignement := FAlignement+'G.0  ---;';
  Ftitre := FTitre + 'Désignation;';
  fLargeur := '5;5;15;70;';
  // --
  laliste := fColnames;
  //
  nbcolsInListe := 0;
  repeat
    lelement := READTOKENST (laliste);
    if lelement <> '' then
    begin
      inc(nbcolsInListe);
    end;
  until lelement = '';

end;

procedure TOF_BTAFFAIREINTERV.LocalDestroy;
begin
  FreeAndNil(fTOBAffaireInterv);
  ImTypeArticle.free;
end;

procedure TOF_BTAFFAIREINTERV.DecodeArgs(S: string);
var LesArgs,UnArg : string;
		UnCode,Unevaleur : string;
    IposEgal : integer;
begin
	lesArgs := S;
  Repeat
  	UnArg := READTOKENST(LesArgs);
    if UnArg <> '' then
    begin
      IposEgal := Pos('=',UnArg);
      if IposEgal > 0 then
      begin
        Uncode := copy(unArg,1,IPosEgal-1);
        UneValeur := copy(unArg,IPosEgal+1,255);
        if UnCode='AFFAIRE' then
        begin
        	fAffaire := UneValeur;
        end
        else if UnCode = 'MODEGESTION' then
        begin
        	fModegestion := UneValeur;
        end
        else if Uncode = 'MANDATAIRE' then
        begin
        	if StrToBool_(UneValeur) then
            fRoleEntreprise := 'MANDATAIRE'
          else
            fRoleEntreprise := 'COTRAITANT';
        end;
      end;
    end;
  Until UnArg='';
  //
  NatureAuxi  := 'FOU';
  //
	if fModegestion='COTRAITANCE' then
  begin
  	if fRoleEntreprise = 'MANDATAIRE' then
    begin
    	fLibTiers := 'Cotraitant';
      fTypeInterv := 'X01';
    end else
    begin
    	fLibTiers := 'Mandataire';
      fTypeInterv := 'X00';
      NatureAuxi  := 'CLI';
    end;
  end else
  begin
    fLibTiers := 'Sous traitant';
    fTypeInterv := 'Y00';
  end;
end;

procedure TOF_BTAFFAIREINTERV.ChargeLesIntervenants;
var QQ      : TQuery;
		indice  : integer;
    req     : string;
    TOBBQE  : TOB;
    CodeTiers,CodeAuxi,RefBancaire : string;
    ZoneTiers : String;
    NumeroRib : integer;
begin

	fMaxInterv := 0;
  //
  SetControltext('TLIB_NUMERORIB', '');
  SetControltext('TLIBREFBANQUE', '');
  //

  if fTypeInterv = 'X00' then
    ZoneTiers := 'BAI_TIERSCLI'
  Else
    ZoneTiers := 'BAI_TIERSFOU';


  Req := 'SELECT BAI_AFFAIRE, BAI_ORDRE,' + ZoneTiers + ' AS BAI_TIERSFOU, BAI_TYPEINTERV, BAI_TYPEPAIE, BAI_NUMEROCONTACT, BAI_NUMERORIB,';

  Req := Req + '(SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=' + ZoneTiers;
  Req := Req + ' AND T_NATUREAUXI="' + NatureAuxi + '") AS TLIBTIERS,';
  Req := Req + '(SELECT R_DOMICILIATION FROM RIB WHERE R_AUXILIAIRE=' + ZoneTiers;
  Req := Req + ' AND R_NUMERORIB=BAI_NUMERORIB) AS TLIB_NUMERORIB,';
  Req := Req + '(SELECT C_NOM FROM CONTACT  WHERE C_TYPECONTACT="T" ';
  Req := Req + ' AND C_NATUREAUXI="'+ NatureAuxi + '" AND C_AUXILIAIRE=';
  Req := Req + '(SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS=' + ZoneTiers + ' AND T_NATUREAUXI="' + NatureAuxi + '") ';

  Req := Req + 'AND C_NUMEROCONTACT=BAI_NUMEROCONTACT)  AS TLIBCONTACT FROM AFFAIREINTERV WHERE BAI_AFFAIRE="'+fAffaire +'"';

  QQ := OpenSql (Req,true,-1,'',true);
  if not QQ.eof then
  begin
  	fTOBAffaireInterv.LoadDetailDB('AFFAIREINTERV','','',QQ,false);
  end;
  ferme (QQ);

  for Indice := 0 to fTOBAffaireInterv.detail.count -1 do
  begin
    //Chargement suite info bancaire...
    TOBBQE := Tob.Create('RIB', nil, -1);

    CodeTiers := fTOBAffaireInterv.Detail[indice].GetValue('BAI_TIERSFOU');
    CodeAuxi  := TiersAuxiliaire (CodeTIers,false,NatureAuxi);

    NumeroRib := StrToInt(fTOBAffaireInterv.Detail[indice].GetValue('BAI_NUMERORIB'));
    LectRib(TOBBQE, CodeAuxi, fTOBAffaireInterv.Detail[indice].GetValue('BAI_NUMERORIB'));
    RefBancaire := '';
		GetCodeBQ(CodeAuxi, NumeroRIB, RefBancaire);
    fTOBAffaireInterv.Detail[indice].putValue('TLIB_NUMERORIB',RefBancaire);

    Charge_InfoBqe(TOBBQE, indice);

    FreeAndNil(TOBBQE);

  	if fMaxInterv < fTOBAffaireInterv.detail[Indice].getValue('BAI_ORDRE') then
			 fMaxInterv := fTOBAffaireInterv.detail[Indice].getValue('BAI_ORDRE');
  end;

end;
Procedure TOF_BTAFFAIREINTERV.Charge_InfoBQE(TOBBQE : TOB; Indice : integer);
begin

    if Assigned(TOBBQE) then
    begin
      fTOBAffaireInterv.Detail[indice].AddChampSupValeur('ETABBQ', TOBBQE.GetString('R_ETABBQ'));
      fTOBAffaireInterv.Detail[indice].AddChampSupValeur('GUICHET',TOBBQE.GetString('R_GUICHET'));
      fTOBAffaireInterv.Detail[indice].AddChampSupValeur('NUMEROCOMPTE', TOBBQE.GetString('R_NUMEROCOMPTE'));
      fTOBAffaireInterv.Detail[indice].AddChampSupValeur('CLERIB', TOBBQE.GetString('R_CLERIB'));
      fTOBAffaireInterv.Detail[indice].AddChampSupValeur('IBAN', TOBBQE.GetString('R_CODEIBAN'));
    end
    else
    begin
      fTOBAffaireInterv.Detail[indice].AddChampSupValeur('ETABBQ', '');
      fTOBAffaireInterv.Detail[indice].AddChampSupValeur('GUICHET', '');
      fTOBAffaireInterv.Detail[indice].AddChampSupValeur('NUMEROCOMPTE', '');
      fTOBAffaireInterv.Detail[indice].AddChampSupValeur('CLERIB', '');
      fTOBAffaireInterv.Detail[indice].AddChampSupValeur('IBAN', '');
    end;

end;


procedure TOF_BTAFFAIREINTERV.ChargeLagrille;
begin
  GS.VidePile(false);
  DefiniGrille;
  if fTOBAffaireInterv.detail.count = 0 then GS.rowCount := 2 else GS.rowCount := fTOBAffaireInterv.detail.count+1;
  AffichelaGrille (false);
  TFVierge(ecran).HMTrad.ResizeGridColumns (GS);
end;


procedure TOF_BTAFFAIREINTERV.AffichelaGrille (Entete : boolean=false);
var Indice : integer;
begin
	for Indice := 0 to fTOBAffaireInterv.detail.Count -1 do
  begin
  	AfficheLaLigne (Indice,Entete);
  end;
  TFVierge(ecran).HMTrad.ResizeGridColumns (GS);
end;

procedure TOF_BTAFFAIREINTERV.AfficheLaLigne (Indice : integer; Entete : boolean=false);
begin
	fTOBAffaireInterv.detail[Indice].PutLigneGrid(gs,Indice+1,Entete,Entete,fColNames);
end;


procedure TOF_BTAFFAIREINTERV.SetEvents(active: boolean);
begin
	if Active then
  begin
		GS.OnrowEnter := GSRowEnter;
		GS.OnDblclick := GSDblclick;
    GS.PostDrawCell := GSPostDrawCell;
    Bdelete.OnClick := BdeleteClick;
    Binsert.OnClick := BInsertClick;
  end else
  begin
		GS.OnrowEnter := nil;
		GS.OnDblclick := nil;
    GS.PostDrawCell := nil;
    Bdelete.OnClick := nil;
    Binsert.OnClick := nil;
  end;
end;

procedure TOF_BTAFFAIREINTERV.Positionne(Arow: integer);
var cancel :boolean;
begin
	GS.row := Arow;
  GSRowEnter(self,GS.row,cancel,false);
end;

procedure TOF_BTAFFAIREINTERV.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin

  TOBL := GetTOBLigne (Ou);

  if TOBL = nil then
  BEGIN
    ResetInfo;
    EXIT;
  END;

  SetInfoligne (TOBL);

  fLastRow := ou;

end;

procedure TOF_BTAFFAIREINTERV.ResetInfo;
begin
	SetControlText ('BAI_TYPEINTERV',fTypeInterv );
  SetControltext('BAI_TIERSCLI','');
  SetControltext('BAI_TIERSFOU','');
  SetControlCaption ('TLIBTIERSCLI','');
  SetControlCaption ('TLIBTIERSFOU','');
  SetControlCaption ('TLIBCONTACT','');
  SetControltext ('BAI_DATECREATION','');
  SetControltext ('BAI_DATEMODIF','');
  SetControltext ('BAI_NUMERORIB','');
  if Assigned(GetControl('TLIB_CODEBQ')) then SetControltext ('TLIB_CODEBQ','');
  SetControltext ('TBAI_CREATEUR','');
  SetControltext ('TBAI_UTILISATEUR','');
end;

procedure TOF_BTAFFAIREINTERV.SetInfoligne(TOBL: TOB);
begin

	TOBL.PutEcran(ecran,PSAISIE);

  if fTypeInterv = 'X00' then
  begin
    BAI_TIERSCLI.Text  := TOBL.GetString('BAI_TIERSFOU');
    SetControlText('TLIBTIERSCLI',TOBL.GetString('TLIBTIERS'));
  end
  else
  begin
    BAI_TIERSFOU.Text  := TOBL.GetString('BAI_TIERSFOU');
    SetControlText('TLIBTIERSFOU',TOBL.GetString('TLIBTIERS'));    
  end;

  SetInfoBqe(TOBL);

  SetControltext('BAI_DATECREATION',DateTimeToStr(TOBL.getValue('BAI_DATECREATION')));
  SetControltext('TBAI_CREATEUR',Rechdom('TTUTILISATEUR',TOBL.getValue('BAI_CREATEUR'),false));
  SetControltext('BAI_DATEMODIF',DateTimeToStr(TOBL.getValue('BAI_DATEMODIF')));
  SetControltext('TBAI_UTILISATEUR',Rechdom('TTUTILISATEUR',TOBL.getValue('BAI_UTILISATEUR'),false));

end;

procedure Tof_BTAFFAIREINTERV.SetInfoBqe(TOBL : TOB);
Begin

  SetControlText('TLIB_NUMERORIB', TOBL.GetString('TLIB_NUMERORIB'));

  //chargement de la reférence bancaire
  if BAI_NUMERORIB.Text <> '0' then
    if (TOBL.GetValue('ETABBQ')='') and (Tobl.getvalue('GUICHET')='') and (TOBL.GetValue('NUMEROCOMPTE')='') and (TOBL.GetValue('CLERIB')='') then
      SetControltext('TLIBREFBANQUE', TOBL.GetValue('IBAN'))
    else
      SetControltext('TLIBREFBANQUE', TOBL.GetValue('ETABBQ') + ' - ' + Tobl.getvalue('GUICHET') + ' - ' + TOBL.GetValue('NUMEROCOMPTE') + '/' + TOBL.GetValue('CLERIB'))
  else
    SetControltext('TLIBREFBANQUE', '');

end;

function TOF_BTAFFAIREINTERV.GetTOBLigne(Arow: integer) : TOB;
begin
	result := nil;
	if Arow > fTOBAffaireInterv.detail.count then exit;
  result := fTOBAffaireInterv.detail[Arow-1];
end;

procedure TOF_BTAFFAIREINTERV.BdeleteClick(Sender: Tobject);
var cancel : boolean;
begin
	SetEvents (false);
  TOBL.DeleteDB(false);
  TOBL.free; TOBL := nil;
  GS.DeleteRow(GS.row);
  if GS.RowCount = 1 then
  BEGIN
    fLastRow := 1;
		ChargeLagrille;
  END;
	SetEvents (true);
  GSRowEnter(Self,GS.row,cancel,false)

end;

procedure TOF_BTAFFAIREINTERV.BinsertClick(Sender: Tobject);
begin

	TOBL := InsertTOBligne;
  SetInfoLigne(TOBL);
  SetEventsSaisie(True);
  SetEvents(false);

  fInserting := true;

end;

procedure TOF_BTAFFAIREINTERV.SetEventsSaisie(Etat: boolean);
begin
	if Etat then
  begin
  	GS.enabled := false;
    BOk.Onclick := BOkClick;
    BAnnul.Onclick := BAnnulClick;
    Bok.visible := true;
		Bvalider.visible := false;
  	BANNUL.visible := true;
  	BFerme.visible := false;
    //
    PSaisie.Enabled := true;
  end else
  begin
  	GS.enabled := true;
    BOk.Onclick := nil;
    Bannul.OnClick := nil;
    Bok.visible := false;
    Bannul.visible := false;
		Bvalider.visible := true;
  	BANNUL.visible := false;
  	BFerme.visible := true;
    //
    PSaisie.Enabled := false;
  end;
end;

procedure TOF_BTAFFAIREINTERV.BOkClick(Sender: Tobject);
var ferr : integer;
    TypeError : string;
begin

  //Appel procedure avant validation pour contrôle des zones saisies
  if not OnBeforeUpdate(TypeError) then
  begin
    if TypeError = 'B' then
      Exit
    else if TypeError = 'N' then
    begin
      if PGIAsk(ErrorMsg, 'Erreur validation') = Mrno then Exit;
    end
    else
    begin
      PGIInfo('Erreur inconnue');
      Exit;
    end;
  end;

	if (fModegestion = 'COTRAITANCE') then
  begin
  	ferr :=  ControleExist (TOBL,BAI_TIERSFOU.text,GetControltext('BAI_TYPEINTERV'));
    if ferr = 1 then
    begin
    	PgiInfo ('impossible : Cet intervenant est déjà défini');
      BAI_TIERSFOU.text := TOBL.getvalue('BAI_TIERSFOU');
      BAI_TIERSFOU.SetFocus;
      exit;
    end;
    if Ferr = 2 then
    begin
    	PgiInfo ('impossible : Le mandataire est déjà défini');
      //BAI_TYPEINTERV.SetFocus;
      exit;
    end;
    //
  end else
  begin
  end;

  if ftypeinterv = 'X00' then
  begin
    if BAI_TIERSCLI.text = '' then
    begin
      BAI_TIERSCLI.SetFocus;
      ErrorMsg := 'Veuillez renseigner le code client.';
      BAI_TIERSCLI.SetFocus;
      exit;
    end
  end
  else
  begin
    if getControltext('BAI_TIERSFOU') = '' then
    begin
      PgiInfo ('Veuillez renseigner le code fournisseur '+fLibTiers);
      BAI_TIERSFOU.SetFocus;
      exit;
    end;
  end;

  TOBL.GetEcran(ecran,PSAISIE);
  // Controle effectuée --> on insere et on raffraichit la grid
  TOBL.InsertOrUpdateDB(false);
  if finserting then
  begin
  	TOBL.ChangeParent(fTOBAffaireInterv,-1);
  	fLastRow := fTOBAffaireInterv.detail.count;
  end;

  SetEventsSaisie(false);
  SetEvents(true);
	ChargeLagrille;
  Positionne(flastRow);
  fInserting := false;
end;

function TOF_BTAFFAIREINTERV.InsertTOBligne: TOB;
begin

	result := TOB.Create('AFFAIREINTERV',nil,-1);

  result.AddChampSupValeur('TLIBTIERSFOU','');
  result.AddChampSupValeur('TLIBTIERSCLI','');
  result.AddChampSupValeur('TLIBTIERS','');
  result.AddChampSupValeur('TLIBCONTACT','');
  result.AddChampSupValeur('TLIB_NUMERORIB','');
  //
  result.AddChampSupValeur('ETABBQ', '');
  result.AddChampSupValeur('GUICHET', '');
  result.AddChampSupValeur('NUMEROCOMPTE', '');
  result.AddChampSupValeur('CLERIB', '00');
  result.AddChampSupValeur('IBAN', '');
  //
  Inc(fMaxInterv);
  //
  result.putValue ('BAI_AFFAIRE',fAffaire);
  result.putValue ('BAI_ORDRE',fMaxInterv);
  result.putValue ('BAI_TYPEINTERV',fTypeInterv);
  result.putvalue ('BAI_DATECREATION',V_PGI.DateEntree);
  result.putvalue ('BAI_CREATEUR',V_PGI.User);
  result.putvalue ('BAI_DATEMODIF',V_PGI.DateEntree);
  result.putvalue ('BAI_UTILISATEUR',V_PGI.User);

end;

procedure TOF_BTAFFAIREINTERV.BAnnulClick(Sender: TObject);
begin
	if fInserting then TOBL.free;
  FExiting := true;
  SetEventsSaisie(false);
  SetEvents(true);
  Positionne(flastRow);
  FExiting := false;
end;

procedure TOF_BTAFFAIREINTERV.ChangeTzTiersSaisie;
var St : string;
    ii : integer;
begin
  St := 'GCTIERSSAISIE';
  ii := TTToNum(St);
  if ii > 0 then
  begin
    if NatureAuxi = 'CLI' then
    	V_PGI.DECombos[ii].Libelle := 'Clients'
    else
      V_PGI.DECombos[ii].Libelle := 'Fournisseurs';
  	V_PGI.DECombos[ii].Where := ' AND (T_NATUREAUXI="'+NatureAuxi+'")';
  end;

end;

procedure TOF_BTAFFAIREINTERV.TIERSCLIEnter(Sender: Tobject);
begin
	ChangeTzTiersSaisie;
end;

procedure TOF_BTAFFAIREINTERV.TIERSCLIEXit(Sender: Tobject);
var LibTiers : string;
begin

	if FExiting then exit;

  if PSaisie.Enabled = False then exit;

  if GetControlText('BAI_TIERSCLI')= '' then Exit;

  if not GetLibTiers(NatureAuxi,BAI_TIERSCLI.text,LibTiers) then
  begin
    PgiInfo ('Ce client n''existe pas. Merci d''en sélectionner un autre');
    if assigned(TOBL) then BAI_TIERSCLI.text := TOBL.GetValue('BAI_TIERSCLI');
    BAI_TIERSCLI.SetFocus;
    exit;
  end
  else
  begin
    SetControlText('TLIBTIERSCLI',LibTiers);
    TOBL.putValue ('TLIBTIERSCLI',LibTiers);
    TOBL.putValue ('TLIBTIERS',LibTiers);
  end;

end;

procedure TOF_BTAFFAIREINTERV.TIERSFOUEnter(Sender: Tobject);
begin
	ChangeTzTiersSaisie;
end;

procedure TOF_BTAFFAIREINTERV.TIERSFOUEXit(Sender: Tobject);
var LibTiers : string;
		ferr     : integer;
    StSQL    : string;
begin

	if FExiting then exit;

	if not GetLibTiers(NatureAuxi,BAI_TIERSFOU.text,LibTiers) then
  begin
  	PgiInfo ('Ce Fournisseur n''existe pas. Merci d''en sélectionner un autre');
    if assigned(TOBL) then BAI_TIERSFOU.text := TOBL.GetValue('BAI_TIERSFOU');
    BAI_TIERSFOU.SetFocus;
    exit;
  end else
  begin
  	ferr :=  ControleExist (TOBL,BAI_TIERSFOU.text,'');
    if ferr = 1 then
    begin
    	PgiInfo ('impossible : Cet intervenant est déjà défini');
      BAI_TIERSFOU.text := TOBL.getvalue('BAI_TIERSFOU');
      BAI_TIERSFOU.SetFocus;
      exit;
    end;
    SetControlText('TLIBTIERSFOU',LibTiers);
    TOBL.putValue('TLIBTIERSFOU',LibTiers);
    TOBL.putValue('TLIBTIERS',LibTiers);
    //recherche des informations bancaires
    StSQL :='SELECT * FROM RIB WHERE R_AUXILIAIRE="' + BAI_TIERSFOU.text + '" AND R_PRINCIPAL="X"';
    //recupération du code banque par défaut....

  end;
end;

procedure TOF_BTAFFAIREINTERV.TIERSCLIElipsisClick(Sender: TObject);
begin

    if GetTiers(BAI_TIERSCLI) then
    begin
      TIERSCLIEXit (self);
    end;

end;

procedure TOF_BTAFFAIREINTERV.TIERSFOUElipsisClick(Sender: TObject);
begin
  if GetTiers(BAI_TIERSFOU) then
  begin
  	TIERSFOUEXit (self);
  end;
end;

procedure TOF_BTAFFAIREINTERV.GSDblclick(Sender: Tobject);
begin

	TOBL := GetTOBLigne (GS.row);

  if assigned(TOBL) then
  begin
    SetInfoLigne(TOBL);
    SetEventsSaisie(True);
    SetEvents(false);
    if BAI_TIERSCLI.visible then BAI_TIERSCLI.SetFocus
    											  else BAI_TIERSFOU.SetFocus;
    fInserting := false;
  end
  else
    BinsertClick(Self);


end;

function TOF_BTAFFAIREINTERV.ControleExist(TOBL: TOB;Fournisseur : string; TypeIntervenant : string): integer;
var indice : integer;
		Sql : string;
begin
	result := 0;
  // Controle dans les cotraitants
	for indice := 0 to fTOBAffaireInterv.detail.count -1 do
  begin
  	if fTOBAffaireInterv.detail[Indice] <> TOBL then
    begin
    	if fTOBAffaireInterv.detail[indice].getvalue('BAI_TIERSFOU') = Fournisseur then
      begin
      	result:= 1;
        break;
      end;
      //
      if (fModegestion = 'COTRAITANCE') and (fRoleEntreprise = 'COTRAITANT') and (TypeIntervenant = 'X00') then
      begin
        if fTOBAffaireInterv.detail[indice].getvalue('BAI_TYPEINTERV') = TypeIntervenant then
        begin
          result:= 2; // il n'y en aura qu'un (de mandataire) !!
          break;
        end;
      end;
    end;
  end;
  // Verification dans tous les documents associés a l'affaire
  Sql := 'SELECT BPI_TIERSFOU FROM PIECEINTERV WHERE '+
  			 'BPI_TIERSFOU="'+Fournisseur+'" AND BPI_NATUREPIECEG="DBT" AND BPI_NUMERO IN '+
         '(SELECT GP_NUMERO FROM PIECE WHERE GP_NATUREPIECEG="DBT" AND GP_AFFAIRE="'+fAffaire+'")';
  if ExisteSql (Sql) then result := 1;
end;

procedure TOF_BTAFFAIREINTERV.DefiniEcran;
begin

	ChargeListImage;

	if fModegestion = 'COTRAITANCE' then
  begin
  	BAI_TYPEINTERV.Plus := ' AND CO_CODE LIKE "X%"';
  	if fRoleEntreprise = 'COTRAITANT' then
    begin
    	//BAI_TYPEINTERV.Enabled := true; // permet de choisir le type d'intervenant à décrire
    end;
  end else
  begin
  	BAI_TYPEINTERV.Plus := ' AND CO_TYPE LIKE "Y%"';
  end;

	if (fModegestion = 'COTRAITANCE') then
  begin
  	if fTypeInterv  = 'X00' then
    begin
    	BAI_TIERSCLI.visible := true;
      GetControl('TBAI_TIERSCLI').Visible := True;
      GetControl('TLIBTIERSCLI').Visible  := True;
      GetControl('TBAI_TIERSFOU').Visible := False;
      GetControl('TLIBTIERSFOU').Visible  := False;
			BAI_TIERSFOU.Visible := false;
    end else
    begin
    	BAI_TIERSCLI.visible := false;
      GetControl('TBAI_TIERSCLI').Visible := False;
      GetControl('TLIBTIERSCLI').Visible  := False;
      GetControl('TBAI_TIERSFOU').Visible := True;
      GetControl('TLIBTIERSFOU').Visible  := True;
			BAI_TIERSFOU.Visible := true;
    end;
  end;


end;

procedure TOF_BTAFFAIREINTERV.GSPostDrawCell(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var ARect: TRect;
  NumGraph : integer;
  fTOBl : TOB;
begin

  if GS.RowHeights[ARow] <= 0 then Exit;

  ARect := GS.CellRect(ACol, ARow);

  GS.Canvas.Pen.Style := psSolid;
  GS.Canvas.Pen.Color := clgray;
  GS.Canvas.Brush.Style := BsSolid;

  if (Acol = 1) and (Arow >= GS.fixedRows) then
  begin
  	fTOBL := GetTOBLigne(Arow); if fTOBL = nil then exit;
    Canvas.FillRect(ARect);
    NumGraph := RecupTypeGraph(fTOBL);
    if NumGraph >= 0 then
    begin
      ImTypeArticle.DrawingStyle := dsnormal;
      ImTypeArticle.Draw(CanVas, ARect.left + round((Arect.Right - ARect.left - ImTypeArticle.Width)/2)  , ARect.top, NumGraph);
    end;
  end;

end;

procedure TOF_BTAFFAIREINTERV.ChargeListImage;
var UneImage : Timage;
    NomImg : string;
    indice : integer;
begin
  Indice := 1;
  NomImg := 'IMG'+IntToStr(Indice);
  UneImage := TImage(GetCOntrol(NomImg));
  repeat
    if UneImage <> nil then
    begin
      if ImTypeArticle.AddMasked  (UneImage.Picture.Bitmap,TColor($FF00FF)) < 0 then exit;
      //
      inc(Indice);
      //
      NomImg := 'IMG'+IntToStr(Indice);
      UneImage := TImage(GetCOntrol(NomImg));
    end;
  until UneImage = nil;
end;

function TOF_BTAFFAIREINTERV.RecupTypeGraph(TOBL: TOB): integer;
var TypeInterv : string;
begin

  TypeInterv := TOBL.GetValue('BAI_TYPEINTERV');

  if TypeInterv = 'X00' Then Result := 0
  else if (TypeInterv = 'X01') then Result := 1
  else if (TypeInterv = 'Y01') then Result := 2
  else Result := -1;

end;

procedure TOF_BTAFFAIREINTERV.ContactClick(Sender: TObject);
var retour,CodeAuxi,CodeTiers,Action : string;
begin

	if fTypeInterv = 'X00' then  // dans le cas de cotraitant
    CodeTiers := BAI_TIERSCLI.Text
  else                         // dans le cas de mandataire
    CodeTiers := BAI_TIERSFOU.Text;

  CodeAuxi    := TiersAuxiliaire(CodeTiers,false,NatureAuxi);

  Action      := 'ACTION=MODIFICATION';
  //
  retour := AglLanceFIche('YY','YYCONTACT','T;'+CodeAuxi,'',Action+ ';TYPE=T;'+'TYPE2='+Natureauxi+';TITRE=Sélection du destinataire;TIERS='+CodeTiers+';ALLCONTACT' );
  if retour <> '' then THEdit(GetControl('BAI_NUMEROCONTACT')).Text := ReadTokenSt(Retour);

end;

procedure TOF_BTAFFAIREINTERV.ContactExit(Sender: tobject);
var libcontact  : String;
    CodeTiers   : String;
    CodeAuxi    : String;
begin

	if FExiting then exit;

  if PSaisie.Enabled = False then exit;
  if StrToInt(BAI_NUMEROCONTACT.text) = 0 then Exit;

	if fTypeInterv = 'X00' then       // dans le cas de cotraitant
    CodeTiers := BAI_TIERSCLI.Text
  else                              // dans le cas de mandataire
    CodeTiers := BAI_TIERSFOU.Text;
  CodeAuxi  := TiersAuxiliaire(CodeTiers,false, NatureAuxi);

	if not GetContact(CodeTiers ,StrToInt(BAI_NUMEROCONTACT.text),Libcontact) then
  begin
  	PgiInfo ('Ce contact n''existe pas. Merci d''en sélectionner un autre');
    SetControlText('BAI_NUMEROCONTACT',TOBL.GetValue('BAI_NUMEROCONTACT'));
    BAI_NUMEROCONTACT.SetFocus;
    exit;
  end else
  begin
    SetControlText('TLIBCONTACT',LibContact);
    TOBL.putValue('TLIBCONTACT',LibContact);
  end;

end;

procedure TOF_BTAFFAIREINTERV.CodeBQOnChange(Sender: TObject);
begin

  if BAI_NUMERORIB.Text <> '' then CodeBQExit (self);

end;

procedure TOF_BTAFFAIREINTERV.CodeBQExit(Sender: TObject);
Var RefBancaire : string;
    CodeTiers   : string;
    CodeAuxi : string;
    NumeroRIB   : Integer;
    TOBBQE      : TOb;
begin

  if FExiting then exit;

  if PSaisie.Enabled = False then exit;
  NumeroRIB := StrToInt(BAI_NUMERORIB.Text);
  if NumeroRIB = 0 then Exit;

	if fTypeInterv = 'X00' then     // dans le cas de cotraitant
    CodeTiers := BAI_TIERSCLI.Text
  else                            // dans le cas de mandataire
    CodeTiers := BAI_TIERSFOU.Text;

  CodeAuxi := TiersAuxiliaire (CodeTIers,false, NATUREAUXI);

  //recherche des informations bancaires
	if not GetCodeBQ(CodeAuxi, NumeroRIB, RefBancaire) then
  begin
  	PgiInfo ('Cette banque n''existe pas. Merci d''en sélectionner une autre');
    BAI_NUMERORIB.Text := TOBL.GetValue('BAI_NUMERORIB');
    exit;
  end
  else
  begin
    //Chargement suite info bancaire...
    TOBBQE := Tob.Create('RIB', nil, -1);
    LectRib(TOBBQE, CodeAuxi, NumeroRIB);
    TOBL.putValue('TLIB_NUMERORIB',RefBancaire);
    if Assigned(TOBBQE) then
    begin
      TOBL.putValue('ETABBQ', TOBBQE.GetString('R_ETABBQ'));
      TOBL.putValue('GUICHET',TOBBQE.GetString('R_GUICHET'));
      TOBL.putValue('NUMEROCOMPTE', TOBBQE.GetString('R_NUMEROCOMPTE'));
      TOBL.putValue('CLERIB', TOBBQE.GetString('R_CLERIB'));
      TOBL.putValue('IBAN', TOBBQE.GetString('R_CODEIBAN'));
    end
    else
    begin
      TOBL.putValue('ETABBQ', '');
      TOBL.putValue('GUICHET', '');
      TOBL.putValue('NUMEROCOMPTE', '');
      TOBL.putValue('CLERIB', '');
      TOBL.putValue('IBAN', '');
    end;
    FreeandNil(TOBBQE);
    SetInfoBqe(TOBL);
  end;

end;


procedure TOF_BTAFFAIREINTERV.RIBElipsisClick(Sender: TObject);
var CodeTiers   : String;
    sNumeroRib  : string;
begin

	if fTypeInterv = 'X00' then
  begin
    // dans le cas de cotraitant
    CodeTiers := BAI_TIERSCLI.Text;
  end else
  begin
    // dans le cas de mandataire
    CodeTiers := BAI_TIERSFOU.Text;
  end;

  //appel de la fiche RIB pour récupération code banque...
  //  sNumeroRIB := AGLLanceFiche('YY','YYRIB',CodeTiers,'','NumAux='+ TiersAuxiliaire (CodeTIers,false,NatureAuxi));
  sNumeroRIB := AGLLanceFiche('YY','YYRIB',TiersAuxiliaire (CodeTIers,false,NatureAuxi),'','NumAux='+ TiersAuxiliaire (CodeTIers,false,NatureAuxi));
  if sNumeroRib <> '' then
  begin
    BAI_NUMERORIB.Text := sNumeroRib;
    CodeBQExit(Self);
  end;

end;

Initialization
  registerclasses ( [ TOF_BTAFFAIREINTERV ] ) ;
end.

