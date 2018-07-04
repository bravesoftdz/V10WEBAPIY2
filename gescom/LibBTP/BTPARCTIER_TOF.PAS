{***********UNITE*************************************************
Auteur  ...... : LS
Créé le ...... : 11/02/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARCTIER ()
Mots clefs ... : TOF;BTPARCTIER
*****************************************************************}
Unit BTPARCTIER_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
{$IFNDEF EAGLCLIENT}
		 fe_main,
     db, uPDFBatch,EdtREtat,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul, MainEagl, UtilEagl,
{$ENDIF}
		 Vierge,	
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOF,
     menus,
     UListByUser,
     BTStructChampSup,
     graphics,
     grids,
     AglInit,
     Windows,
     CalcolegenericBTP,
     UserialiseLSE,
     UtilParc ;

Type
	TmemoClef = class (TObject)
  	private
  		N1,N2,N3,N4,N5 : integer;
  	public
    	constructor create;
      destructor destroy; override;
      procedure InitMemo;
      procedure MemoriseClef (TOBL: TOB);
      function RetrouvePosition (TOBLigne : TOB) : integer;
  end;

  TOF_BTPARCTIER = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	freqPere : string;
  	fniveau : integer;
  	fmemoKey : TmemoClef;
  	fTOBparc : TOB;
    fTOBPere : TOB;
    fAction : TActionFiche;
    fSelectSql,fOrderBy : string;
    fListe,fColonnes : string;
    fCaption : string;
    FNomTable, FLien, FSortBy, stCols : string;
    Ftitre : Hstring;
    FLargeur, FAlignement, FParams, FPerso: string;
    title, NC : Hstring;
    OkTri, OkNumCol : boolean;
    nbcolsinliste : integer;

    // ZONES DE LA FICHE
    GS : THgrid;
    ELTPARC,LIBELLEPARC : Thedit;
    TYPEARTICLE : THMultiValComboBox;
    FAMILLENIV1,FAMILLENIV2,FAMILLENIV3 : THValComboBox;
    SOUSCONTRAT,TOUTPARC : THCheckbox;
    XX_N1,XX_N2,XX_N3,XX_N4,XX_N5 : THEDIT;
    CODEINTERNEP,TIERS,XX_WHERE,CODEARTICLEP : THEdit;
    TBSELECT,BPARAMLISTE,BSELECTALL,BINSMULTIPLE : TToolbarButton97;
    POPACTIONG : TPopupMenu;
    PC1 : TPageControl;
    BDELETE,BImprimer : TToolbarButton97;

    MnVoirDetail,MnVoirElt,MnAffecteVers,MnAffecteContrat : TmenuItem;
    procedure DefiniCaption;
    procedure CreateTOB;
    procedure GetComponents;
  	procedure SetArguments (S : string);
    procedure SetEvents;
    procedure InitEntete;
    procedure GSDBlClick (Sender : TObject);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GSDEVFlipSelection(Sender: TObject);
    procedure BPARAMListeClick (Sender : Tobject);
    procedure SelectClick (Sender : Tobject);
    function GetTOBPARC (ligne : integer) : TOB;
    function ConstitueArg (TOBL : TOB; ActionSuiv : TActionFiche;FromLigne:boolean=false) : string;
    function ConstitueRequete : string;
    function DecodeTypeArticle (TypeArticle : string) : string;
    procedure positionne (ligne : integer);
    procedure DefiniGrille;
    procedure SetparamGrille;
    procedure AffichelaGrille(Entete: boolean=false);
    procedure AfficheLaLigne(Indice: integer; Entete: boolean=false);
    function RemplaceLibelleFamilles(stChamp: string): string;
    procedure VoirDetail(Sender: TObject);
    procedure VoireElt (Sender : TObject);
    function IsOneSelected : boolean;
    procedure SelectionTout (sender : Tobject);
    procedure positionneMenus;
    procedure ChgVersions (sender : TObject);
    procedure AfficheTout;
    procedure Affecteversion (TOBL : TOB; Version : string);
    procedure SetContrat (Sender : TObject);
    procedure AffecteContrat(TOBL: TOB; Contrat: string);
    procedure BdeleteClick (Sender : TOBject);
    function ConstitueReqDelete (TOBL : TOB) : string;
    procedure Bimprimerclick (Sender : TObject);
    procedure BinsMultipleClick(Sender : Tobject);
    procedure AjouteElementsParcs(UneTOBArt : TOB);
    function ConstitueRequetePere: string;
    procedure ChargeTOBpere;
    function FindMaxIndice: integer;

end ;

Implementation
uses UtilArticle,AglInitGc;
procedure TOF_BTPARCTIER.OnNew ;
var stargument : string;
		TOBL : TOB;
    Arow : integer;
begin
  Inherited ;
	TOBL := GetTOBPARC (GS.row);
  if TOBL <> nil then  fmemoKey.MemoriseClef (TOBL);
  StArgument := ConstitueArg (nil,taCreat);
  //
  AGLLancefiche ('BTP','BTPARCELT','','',StArgument);
  SelectClick (self);
  Arow := fmemokey.RetrouvePosition (fTOBparc);
  if Arow < 0 then
  begin
  	Arow := 1
  end else inc(Arow);
  positionne (Arow);
end ;

procedure TOF_BTPARCTIER.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARCTIER.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTPARCTIER.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPARCTIER.OnArgument (S : String ) ;
begin
  Inherited ;
  fListe := 'BPARCTIERS';
  fSelectSql := 'SELECT * FROM BTPARCTIERS';
	fOrderBy := 'BP1_TIERS,BP1_N1,BP1_N2,BP1_N3,BP1_N4,BP1_N5';
  CreateTOB;
  GetComponents;
  InitEntete;
  SetArguments (S);
  freqPere := ConstitueRequetePere;
  ChargeTOBpere;
  DefiniCaption;
  Ecran.Caption := fCaption; UpdateCaption(ecran);
  SetEvents;
  PC1.ActivePageIndex := 0;
  SetparamGrille;
  SelectClick(Self);
  positionneMenus;

end ;

procedure TOF_BTPARCTIER.OnClose ;
begin
	FreeAndNil(fmemoKey);
	FreeAndNil(fTOBparc);
  if fTOBPere <> nil then FreeAndNil(fTOBPere);
  Inherited ;
end ;

procedure TOF_BTPARCTIER.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARCTIER.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARCTIER.CreateTOB;
begin
	fmemoKey := TmemoClef.create;
	fTOBparc := TOB.create ('PARCTIERS',nil,-1);
  fTOBPere := TOB.Create ('PARCTIERS',nil,-1);
end;

procedure TOF_BTPARCTIER.GetComponents;
begin
  GS := THgrid (GetControl('GS'));
  ELTPARC := THEdit (GetControl('ELTPARC'));
  LIBELLEPARC := Thedit (GetControl('LIBELLEPARC'));
  TYPEARTICLE := THMultiValComboBox (GetControl('TYPEARTICLE'));
  FAMILLENIV1 := THValComboBox (getControl('FAMILLENIV1'));
  FAMILLENIV2 := THValComboBox (getControl('FAMILLENIV2'));
  FAMILLENIV3 := THValComboBox (GetControl('FAMILLENIV3'));
  SOUSCONTRAT := THCHECKBOX(GetControl('SOUSCONTRAT'));
  TOUTPARC := THCheckbox (GetControl('TOUTPARC'));
  XX_N1 := THEdit (GetControl('XX_N1'));
  XX_N2 := THEDit(getControl('XX_N2'));
  XX_N3 := THEDit(getControl('XX_N3'));
  XX_N4 := THEDit(getControl('XX_N4'));
  XX_N5 := THEDit(getControl('XX_N5'));
  CODEINTERNEP := THEDIt(GetCOntrol('CODEINTERNEP'));
  CODEARTICLEP := THEDIt(GetCOntrol('CODEARTICLEP'));
  TIERS := THEdit (GetControl('TIERS'));
  XX_WHERE := THEdit (GetControl('XX_WHERE'));
  TBSELECT := TToolbarButton97(GEtControl('TBSELECT'));
  POPACTIONG := TPopupMenu (GetControl('POPACTIONG'));
  PC1 := TPageControl(GetControl ('PC1'));
  BPARAMLISTE := TToolbarButton97 (GetControl('BPARAMLISTE'));
  MnVoirDetail := TmenuItem(GetControl('MnVoirDet'));
  MnVoirElt := TmenuItem(GetControl('MnVoirElt'));
  MnAffecteVers:= TmenuItem(GetControl('MnAffecteVers'));
  MnAffecteContrat:= TmenuItem(GetControl('MnAffecteContrat'));
  BSELECTALL := TToolbarButton97 (GetControl('BSELECTALL'));
  BinsMultiple := TToolbarButton97 (GetControl('BINSMULTIPLE'));
  BDELETE := TToolbarButton97(getControl('BDELETE'));
  BImprimer := TToolbarButton97(GetControl('BImprimer'));

end;

procedure TOF_BTPARCTIER.SetArguments(S: string);
var Theparams : string;
		TheParam,TheChamps,TheValue : string;
begin
  TheParams := S;
  XX_N1.text := '0';
  XX_N2.text := '0';
  XX_N3.text := '0';
  XX_N4.text := '0';
  XX_N5.text := '0';
  repeat
  	TheParam := READTOKENST(TheParams);
    if TheParam = '' then break;
    theChamps := READTOKENPipe (TheParam,'=');
    thevalue := TheParam;
    if TheChamps = 'TIERS' then
    begin
    	TIERS.Text := Thevalue;
    end else if TheChamps = 'CODEINTERNEP' then
    begin
    	CODEINTERNEP.Text := TheValue;
    end else if TheChamps = 'CODEARTICLEP' then
    begin
    	CODEARTICLEP.Text := TheValue;
    end else if TheChamps = 'N1' then
    begin
    	if (strtoint(thevalue) > 0) then if fniveau < 1 then fniveau := 1;
	    XX_N1.text := InttoStr(StrToInt(TheValue));
    end else if TheChamps = 'N2' then
    begin
    	if (strtoint(thevalue) > 0) then if fniveau < 2 then fniveau := 2;
	    XX_N2.text := InttoStr(StrToInt(TheValue));
    end else if TheChamps = 'N3' then
    begin
    	if (strtoint(thevalue) > 0) then if fniveau < 3 then fniveau := 3;
	    XX_N3.text := InttoStr(StrToInt(TheValue));
    end else if theChamps = 'N4' then
    begin
    	if (strtoint(thevalue) > 0) then if fniveau < 4 then fniveau := 4;
	    XX_N4.text := InttoStr(StrToInt(TheValue));
    end else if theChamps = 'N5' then
    begin
    	if (strtoint(thevalue) > 0) then if fniveau < 4 then fniveau := 5;
	    XX_N5.text := InttoStr(StrToInt(TheValue));
    end else if theChamps = 'XX_XHERE' then
    begin
    	XX_WHERE.Text := TheValue;
    end else if TheChamps = 'ACTION' then
    begin
    	if theValue = 'CONSULTATION' then fAction := taConsult else fAction := Tamodif;
    end;
  until TheParams = '';
end;

procedure TOF_BTPARCTIER.DefiniCaption;
begin
  if (CODEARTICLEP.text='') then
  begin
    fCaption := 'Eléments du parc client : '+TIERS.Text;
  end else
  begin
    fCaption := 'Détail de la nomenclature '+CODEARTICLEP.text+' du parc client : '+TIERS.Text;
  end;
end;

procedure TOF_BTPARCTIER.SetEvents;
begin
  GS.OnDblClick := GSDBlClick;
  GS.OnRowEnter := GSRowEnter;
  GS.GetCellCanvas := GetCellCanvas;
  GS.OnFlipSelection  := GSDEVFlipSelection;
  TBSELECT.OnClick := SelectClick;
  BPARAMLISTE.OnClick := BPARAMListeClick;
  MnVoirDetail.OnClick := VoirDetail;
  MnVoirElt.OnClick := VoireElt;
  BSELECTALL.OnClick := SelectionTout;
  MnAffecteVers.OnClick := ChgVersions;
  MnAffecteContrat.onclick := SetContrat;
  BDELETE.onclick := BdeleteClick;
  BImprimer.onclick := Bimprimerclick;
  BinsMultiple.onclick := BinsMultipleClick;
end;

procedure TOF_BTPARCTIER.InitEntete;
begin
  ELTPARC.Text := '';
  LIBELLEPARC.Text := '';
  CODEINTERNEP.text := '';
  CODEARTICLEP.text := '';
  TIERS.text := '';
  XX_WHERE.text := '';
end;

procedure TOF_BTPARCTIER.GSDBlClick(Sender: TObject);
var stArgument : string;
		TOBL : TOB;
    Arow : integer;
begin
	TOBL := GetTOBPARC (GS.row);  if TOBL = nil then exit;
  fmemoKey.MemoriseClef (TOBL);
  StArgument := ConstitueArg (TOBL,fAction);
  //
  AGLLancefiche ('BTP','BTPARCELT','','',StArgument);
  //
  SelectClick (self);
  Arow := fmemokey.RetrouvePosition (fTOBparc);
  if Arow < 0 then
  begin
  	Arow := 1
  end else inc(Arow);
  positionne (Arow);
end;

procedure TOF_BTPARCTIER.AfficheTout;
begin
  GS.VidePile(false);
  DefiniGrille;
  if fTOBparc.detail.count = 0 then GS.rowCount := 2 else GS.rowCount := fTOBparc.detail.count+1;
  AffichelaGrille (false);
  TFVierge(ecran).HMTrad.ResizeGridColumns (GS);
end;

procedure TOF_BTPARCTIER.SelectClick(Sender: Tobject);
var req : string;
		QQ : Tquery;
begin
	fTOBparc.ClearDetail;
	Req := fSelectSql+' WHERE '+ConstitueRequete+' ORDER BY '+fOrderBy;
	MemoriseChampsSupPARCL (req);
  QQ := OpenSql (Req,true,-1,'',true);
  if not QQ.eof then fTOBparc.LoadDetailDB('PARCTIERS','','',QQ,false);
  AfficheTout;
  positionneMenus;
  positionne(1);
end;

function TOF_BTPARCTIER.GetTOBPARC(ligne: integer): TOB;
begin
	result := nil;
	if Ligne > fTOBparc.detail.count then exit;
  result := fTOBparc.detail[Ligne-1];
end;

function TOF_BTPARCTIER.ConstitueArg(TOBL: TOB;ActionSuiv : TActionFiche;FromLigne:boolean=false): string;
begin
	if TOBL = nil then
  begin
    result := 'TIERS='+TIERS.Text+';CODEARTICLEP='+CODEARTICLEP.text+';CODEINTERNEP='+CODEINTERNEP.text+
              ';N5='+XX_N5.text+
              ';N4='+XX_N4.text+
              ';N3='+XX_N3.text+
              ';N2='+XX_N2.text+
              ';N1='+XX_N1.text;

  end else
  begin
  	if not FromLigne then
    begin
      result := 'TIERS='+TOBL.getValue('BP1_TIERS')+';CODEARTICLEP='+TOBL.GetValue('BP1_CODEARTICLE')+
                ';CODEINTERNEP='+CODEINTERNEP.text;
    end else
    begin
			result := 'TIERS='+TOBL.getValue('BP1_TIERS')+';CODEARTICLEP='+TOBL.GetValue('BP1_CODEARTICLE')+
                ';CODEINTERNEP='+TOBL.GetValue('BP1_CODEINTERNE');
    end;
    if TOBL.getValue('BP1_N5')> 0 then
    begin
      result := result +
                ';N5='+IntToStr(TOBL.GetValue('BP1_N5'))+
                ';N4='+IntToStr(TOBL.GetValue('BP1_N4'))+
                ';N3='+IntToStr(TOBL.GetValue('BP1_N3'))+
                ';N2='+IntToStr(TOBL.GetValue('BP1_N2'))+
                ';N1='+IntToStr(TOBL.GetValue('BP1_N1'));
    end else if TOBL.getValue('BP1_N4')> 0 then
    begin
      result := result +
                ';N4='+IntToStr(TOBL.GetValue('BP1_N4'))+
                ';N3='+IntToStr(TOBL.GetValue('BP1_N3'))+
                ';N2='+IntToStr(TOBL.GetValue('BP1_N2'))+
                ';N1='+IntToStr(TOBL.GetValue('BP1_N1'));
    end else if TOBL.getValue('BP1_N3')> 0 then
    begin
      result := result +
                ';N3='+IntToStr(TOBL.GetValue('BP1_N3'))+
                ';N2='+IntToStr(TOBL.GetValue('BP1_N2'))+
                ';N1='+IntToStr(TOBL.GetValue('BP1_N1'));
    end else if TOBL.getValue('BP1_N2')> 0 then
    begin
      result := result +
                ';N2='+IntToStr(TOBL.GetValue('BP1_N2'))+
                ';N1='+IntToStr(TOBL.GetValue('BP1_N1'));
    end else if TOBL.getValue('BP1_N1')> 0 then
    begin
      result := result +
                ';N1='+IntToStr(TOBL.GetValue('BP1_N1'));
    end;
  end;
  if ActionSuiv=TaConsult then
  	result := result +';ACTION=CONSULTATION'
  else if ActionSuiv = Tamodif then
  	result := result +';ACTION=MODIFICATION'
  else result := result +';ACTION=CREATION';
end;

function TOF_BTPARCTIER.ConstitueRequete: string;
begin
	result := 'BP1_TIERS="'+TIERS.text+'"';
  //
  if ELTPARC.text <> '' then
  begin
  	result := result + ' AND BP1_CODEARTICLE LIKE "'+ELTPARC.text+'%"';
  end;
  if LIBELLEPARC.text <> '' then
  begin
  	result := result + ' AND GA_LIBELLE LIKE "'+LIBELLEPARC.text+'%"';
  end;
  if TYPEARTICLE.VALUE <> '' then
  begin
  	result := result + ' AND BP1_TYPEARTICLE IN ('+DecodeTypeArticle	(TYPEARTICLE.Value)+')';
  end;
  if FAMILLENIV1.Value <> '' then
  begin
  	result := result + ' AND GA_FAMILLENIV1 = "'+FAMILLENIV1.Value+'"';
  end;
  if FAMILLENIV2.Value <> '' then
  begin
  	result := result + ' AND GA_FAMILLENIV2 = "'+FAMILLENIV2.Value+'"';
  end;
  if FAMILLENIV3.Value <> '' then
  begin
  	result := result + ' AND GA_FAMILLENIV3 = "'+FAMILLENIV3.Value+'"';
  end;
  //
  if valeur(XX_N5.text) > 0 then
  begin
  	result := result +
    					' AND BP1_N5='+XX_N5.text+
    					' AND BP1_N4='+XX_N4.text+
    					' AND BP1_N3='+ XX_N3.text+
    					' AND BP1_N2='+ XX_N2.text+
    					' AND BP1_N1='+ XX_N1.text;
  end else if valeur(XX_N4.text) > 0 then
  begin
  	result := result + ' AND BP1_N4='+XX_N4.text+
    					' AND BP1_N3='+ XX_N3.text+
    					' AND BP1_N2='+ XX_N2.text+
    					' AND BP1_N1='+ XX_N1.text;
    result := result + ' AND BP1_N5 <> 0';
  end else if valeur(XX_N3.text) > 0 then
  begin
  	result := result + ' AND BP1_N3='+ XX_N3.text+
    					' AND BP1_N2='+ XX_N2.text+
    					' AND BP1_N1='+ XX_N1.text;
    if not TOUTPARC.Checked then result := result + 'AND BP1_N4<>0 AND BP1_N5=0';
  end else if valeur(XX_N2.text) > 0 then
  begin
  	result := result + ' AND BP1_N2='+ XX_N2.text+
    					' AND BP1_N1='+ XX_N1.text;
    if not TOUTPARC.Checked then result := ' AND BP1_N3<>0 AND BP1_N4=0';
  end else if valeur(XX_N1.text) > 0 then
  begin
  	result := result + ' AND BP1_N1='+ XX_N1.text;
    if not TOUTPARC.Checked then result := result + ' AND BP1_N2<> 0 AND BP1_N3=0';
  end else
  begin
  	if not TOUTPARC.checked then result := result + ' AND BP1_N2=0';
  end;
end;

function TOF_BTPARCTIER.DecodeTypeArticle(TypeArticle: string): string;
var lesvaleurs,OneValeur : string;
		first : boolean;
begin
	result := '';
  lesValeurs := typeArticle;
  first := true;
  repeat
  	OneValeur := READTOKENST(Lesvaleurs);
    if OneValeur <> '' then
    begin
      if not first then result := result +',';
      result := result + '"'+OneValeur+'"';
      if first then first := false;
    end;
  until OneValeur ='';
end;

procedure TOF_BTPARCTIER.AffichelaGrille (Entete : boolean=false);
var Indice : integer;
begin
	for Indice := 0 to fTOBparc.detail.Count -1 do
  begin
  	AfficheLaLigne (Indice,Entete);
  end;
  TFVierge(ecran).HMTrad.ResizeGridColumns (GS);
end;

procedure TOF_BTPARCTIER.AfficheLaLigne (Indice : integer; Entete : boolean=false);
begin
	fTOBparc.detail[Indice].PutLigneGrid(gs,Indice+1,Entete,Entete,fColonnes);
end;

procedure TOF_BTPARCTIER.BPARAMListeClick(Sender: Tobject);
begin
//    doParamListe();
end;

procedure TOF_BTPARCTIER.positionne(ligne: integer);
var Cancel,chg : boolean;
begin
	cancel := false;
	if Ligne <= gs.RowCount then GS.row := Ligne;
  GSRowEnter(self,GS.row,cancel,chg)
end;

procedure TOF_BTPARCTIER.SetparamGrille;
var lelement,laliste : string;
begin
  // récupération du paramétrage général des grilles

  ChargeHListe (fListe, FNomTable, FLien, FSortBy, stCols, FTitre,
                FLargeur, FAlignement, FParams, title, NC, FPerso, OkTri, OkNumCol);
  fColonnes  := stCols;
  laliste := stcols;
  nbcolsInListe := 0;
  repeat
    lelement := READTOKENST (laliste);
    if lelement <> '' then
    begin
      inc(nbcolsInListe);
    end;
  until lelement = '';

end;


procedure TOF_BTPARCTIER.DefiniGrille;
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre,TheLibelle : string;
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

  st := fColonnes;
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
    Thelibelle := RemplaceLibelleFamilles (Nam);
    GS.cells[Indice,0] := leTitre;
    GS.ColNames [Indice] := Nam;
    if copy(Alignement,1,1)='G' then GS.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then GS.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then GS.ColAligns[indice] := taCenter;
    okImg := (copy(Alignement,8,1)='X');
    if OkVisu then
    begin
    	GS.ColWidths[indice] := strtoint(lalargeur);
    end else
    begin
    	GS.ColWidths[indice] := -1;
    end;

    if (OkLib) or (okImg) then
    begin
    	GS.ColFormats[indice] := 'CB=' + Get_Join(Nam);
      if OkImg then GS.ColDrawingModes[Indice]:= 'IMAGE';
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
    if nam = 'BP1_ISCONTRAT' then
    begin
      GS.ColTypes [indice] := 'B' ;
      GS.colaligns[indice]:= tacenter;
      GS.colformats[indice]:= inttostr(Integer(csCoche));
    end;
  end ;
end;

function TOF_BTPARCTIER.RemplaceLibelleFamilles (stChamp : string): string;
begin
	result := '';
  if copy(UpperCase (stChamp),1,7)='LIBPARC' then
  begin
    result := RechDom('BTLIBARTPARC','PA'+Copy(stChamp,8,1),false);
  end;
end;

{ TmemoClef }

constructor TmemoClef.create;
begin
  initMemo;
end;

destructor TmemoClef.destroy;
begin

  inherited;
end;

procedure TmemoClef.InitMemo;
begin
	N1 := 0;
  N2 := 0;
  N3 := 0;
  N4 := 0;
  N5 := 0;
end;

procedure TmemoClef.MemoriseClef(TOBL: TOB);
begin
	if TOBL = nil then exit;
  N1 := TOBL.getValue('BP1_N1');
  N2 := TOBL.getValue('BP1_N2');
  N3 := TOBL.getValue('BP1_N3');
  N4 := TOBL.getValue('BP1_N4');
  N5 := TOBL.getValue('BP1_N5');
end;

function TmemoClef.RetrouvePosition(TOBLigne: TOB): integer;
var Indice : integer;
		found : boolean;
begin
	result := -1;
	found := false;
  For indice := 0 to TOBLigne.detail.count -1 do
  begin
  	if (TOBLigne.detail[Indice].getValue('BP1_N1')=N1) and
    	 (TOBLigne.detail[Indice].getValue('BP1_N2')=N2) and
			 (TOBLigne.detail[Indice].getValue('BP1_N3')=N3) and
			 (TOBLigne.detail[Indice].getValue('BP1_N4')=N4) and
			 (TOBLigne.detail[Indice].getValue('BP1_N5')=N5) then
    begin
    	found := true;
      break;
    end;
  end;
  if found then result := indice;
  InitMemo;
end;

procedure TOF_BTPARCTIER.GSRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
var TOBL : TOB;
begin
	MnVoirDetail.Enabled := false;
  MnVoirElt.Enabled := false;
	TOBL := GetTOBPARC (Ou);  if TOBL = nil then exit;
  MnVoirElt.Enabled := true;
  if TOBL.GetValue('BP1_TYPEARTICLE')='PA2' then MnVoirDetail.Enabled := true;
end;


procedure TOF_BTPARCTIER.VoirDetail(Sender: TObject);
var stArgument : string;
		TOBL : TOB;
begin
	TOBL := GetTOBPARC (GS.row);  if TOBL = nil then exit;
  StArgument := ConstitueArg (TOBL,fAction,true);
  //
  AGLLancefiche ('BTP','BTPARCTIER','','',StArgument);
  //
end;

procedure TOF_BTPARCTIER.VoireElt(Sender: TObject);
begin
	GSDBlClick (self);	
end;

procedure TOF_BTPARCTIER.GetCellCanvas(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var TOBL : TOB;
begin
  if ACol < GS.FixedCols then Exit;
  if Arow < GS.Fixedrows then Exit;
	TOBL := GetTOBPARC (Arow);  if TOBL = nil then exit;
  if (TOBL.GetValue('DATEFINCONTRAT')<Date) or (TOBL.GetValue('DATERESILIATION')<Date) or (TOBL.GetValue('ETATAFFAIRE')='TER') then
  begin
  	if (gdSelected in Astate) then
    begin
    	Canvas.Brush.Color := Clred;
      Canvas.Font.Color := clCaptionText;
      Canvas.Font.Style := Canvas.Font.Style + [fsbold];
    end else
    begin
      Canvas.Font.Color := clRed;
      Canvas.Font.Style := Canvas.Font.Style + [fsbold];
    end;
  end;
  if (TOBL.GetValue('BP1_ETATPARC')='HS') then
  begin
    Canvas.Font.Style := Canvas.Font.Style + [fsStrikeOut];
  end;

end;
procedure TOF_BTPARCTIER.positionneMenus;
begin
  if (GS.AllSelected) or (IsOneSelected) then
  begin
    MnAffecteVers.enabled := true;
    MnAffecteContrat.enabled := true;
  end else
  begin
    MnAffecteVers.enabled := false;
    MnAffecteContrat.enabled := false;
  end;
end;

procedure TOF_BTPARCTIER.GSDEVFlipSelection(Sender: TObject);
begin
	positionneMenus;
end;

function TOF_BTPARCTIER.IsOneSelected: boolean;
var Indice : integer;
begin
	result := false;
	for Indice := 1 to GS.rowcount-1 do
  begin
  	if GS.IsSelected(Indice) then
    begin
    	result := true;
      break;
    end;
  end;
end;

procedure TOF_BTPARCTIER.SelectionTout (sender : Tobject);
begin
	if GS.AllSelected then GS.AllSelected := false else GS.AllSelected := true;
  positionneMenus;
end;

procedure TOF_BTPARCTIER.ChgVersions(sender: TObject);
var CodeArticles,version : string;
		Indice : integer;
    first : boolean;
    TOBRET,TOBL : TOB;
begin
	CodeArticles := '';
  first := true;
  For Indice := 1 to GS.rowcount-1 do
  begin
		TOBL := GetTOBPARC (indice);  if TOBL = nil then continue;
  	if GS.IsSelected(Indice) or (GS.AllSelected) then
    begin
    	if first then
      begin
      	first := false;
      end else
      begin
      	CodeArticles:=CodeArticles+';';
      end;
      CodeArticles := CodeArticles + TOBL.GetValue('BP1_CODEARTICLE');
    end;
  end;
  TOBRet := TOB.Create ('UN ECHANGE',nil,-1);
  TOBREt.AddChampSupValeur('ARTICLES',CodeArticles);
  TOBREt.AddChampSupValeur ('FORCE','-');
  TOBREt.AddChampSupValeur('RETOUR','');
  TheTOB := TOBret;
  AGLLanceFiche('BTP','BTCHANGEVERS','','','ACTION=MODIFICATION');
  TheTOB := nil;
  version := TOBret.getValue('RETOUR');
  if version <>'' then
  begin
    For Indice := 1 to GS.rowcount-1 do
    begin
      if GS.IsSelected(Indice) or (GS.AllSelected) then
      begin
      	TOBL := GetTOBPARC (indice);  if TOBL = nil then continue;
        Affecteversion(TOBL,Version);
      end;
    end;
  end;
  GS.ClearSelected;
  AfficheTout;
  TOBRET.free;
end;

procedure TOF_BTPARCTIER.Affecteversion(TOBL: TOB; Version : string);
var CodeSeria : string;
begin
	CodeSeria := GetSeriaLSE(CODEINTERNEP.Text,TOBL.getValue('BP1_CODEINTERNE'),version,TOBL.GEtValue('BP1_QTE'),TOBL.GEtValue('BP1_DATEFINSERIA'));
  TOBL.putValue('BP1_CODEVERSION',version);
  if CodeSeria <> '' then TOBL.putValue('BP1_CODEDEBLOC',CodeSeria);
  TOBL.UpdateDB(false);
  TOBL.SetAllModifie(false);
end;

procedure TOF_BTPARCTIER.AffecteContrat(TOBL : TOB; Contrat: string);
var TheContrat,Aff0,Aff1,Aff2,Aff3,Avenant : string;
begin
	theCOntrat := contrat;
	TheContrat := READTOKENST(TheContrat);
	BTPCodeAffaireDecoupe (Contrat,aff0,aff1,aff2,aff3,avenant,tacreat,true);
  TOBL.putValue('BP1_AFFAIRE',TheContrat);
  TOBL.putValue('BP1_AFFAIRE1',aff1);
  TOBL.putValue('BP1_AFFAIRE2',aff2);
  TOBL.putValue('BP1_AFFAIRE3',aff3);
  TOBL.putvalue('BP1_ISCONTRAT','X');
  TOBL.UpdateDB(false);
  TOBL.SetAllModifie(false);
end;

procedure TOF_BTPARCTIER.SetContrat(Sender: TObject);
var Indice : integer;
    Contrat : string;
    Okreaffiche : boolean;
    TOBL : TOB;
begin
	Okreaffiche := false;
  Contrat := AglLanceFiche('BTP','BTAFFAIRE_MUL','AFF_AFFAIRE0=I','','AFF_TIERS=' +TIERS.Text + ';ETAT=ENC;STATUT=INT;ACTION=RECH');   // Affaires
  if Contrat <> '' then
  begin
    For Indice := 1 to GS.rowcount-1 do
    begin
      if GS.IsSelected(Indice) or (GS.AllSelected) then
      begin
      	TOBL := GetTOBPARC (indice);  if TOBL = nil then continue;
        AffecteContrat(TOBL,Contrat);
      end;
    end;
  end;
  GS.ClearSelected;
  SelectClick (self);
end;

procedure TOF_BTPARCTIER.BdeleteClick(Sender: TOBject);
var REq : string;
		TOBL : TOB;
begin
	if PGiAsk('Vous êtes sur le point de supprimer cet élément. Confirmez-vous ?')=Mryes then
  begin
		TOBL := GetTOBPARC (GS.row);  if TOBL = nil then exit;
    req := ConstitueReqDelete (TOBL);
    ExecuteSql ('DELETE FROM PARCTIERS WHERE '+req);
    GS.ClearSelected;
    SelectClick (self);
  end;

end;

function TOF_BTPARCTIER.ConstitueReqDelete(TOBL: TOB): string;
begin
	result := 'BP1_TIERS="'+TIERS.text+'"';
  //
  if TOBL.getValue('BP1_N5') > 0 then
  begin
  	result := result +
    					' AND BP1_N5='+ inttostr(TOBL.getValue('BP1_N5'))+
    					' AND BP1_N4='+ inttostr(TOBL.getValue('BP1_N4'))+
    					' AND BP1_N3='+ inttostr(TOBL.getValue('BP1_N3'))+
    					' AND BP1_N2='+ inttostr(TOBL.getValue('BP1_N2'))+
    					' AND BP1_N1='+ inttostr(TOBL.getValue('BP1_N1'));
  end else if TOBL.getValue('BP1_N4') > 0 then
  begin
  	result := result +
    					' AND BP1_N4='+ inttostr(TOBL.getValue('BP1_N4'))+
    					' AND BP1_N3='+ inttostr(TOBL.getValue('BP1_N3'))+
    					' AND BP1_N2='+ inttostr(TOBL.getValue('BP1_N2'))+
    					' AND BP1_N1='+ inttostr(TOBL.getValue('BP1_N1'));
  end else if TOBL.getValue('BP1_N3') > 0 then
  begin
  	result := result +
    					' AND BP1_N3='+ inttostr(TOBL.getValue('BP1_N3'))+
    					' AND BP1_N2='+ inttostr(TOBL.getValue('BP1_N2'))+
    					' AND BP1_N1='+ inttostr(TOBL.getValue('BP1_N1'));
  end else if TOBL.getValue('BP1_N2') > 0 then
  begin
  	result := result +
    					' AND BP1_N2='+ inttostr(TOBL.getValue('BP1_N2'))+
    					' AND BP1_N1='+ inttostr(TOBL.getValue('BP1_N1'));
  end else if TOBL.getValue('BP1_N1') > 0 then
  begin
  	result := result +
    					' AND BP1_N1='+ inttostr(TOBL.getValue('BP1_N1'));
  end;

end;

procedure TOF_BTPARCTIER.Bimprimerclick(Sender: TObject);
var TOBEdt,TOBEdit,TOBCLi,TOBP : TOB;
		Indice : integer;
		QQ : TQuery;
begin
  TOBEdit := TOB.Create ('LA TOB EDT',nil,-1);
  TOBEdt := TOB.Create ('TOB EDT',TOBEdit,-1);
  TOBCLi := TOB.Create('TIERS',nil,-1);
  //
  QQ := OpenSql ('SELECT * FROM TIERS WHERE T_TIERS="'+Tiers.text+'" AND T_NATUREAUXI="CLI"',true,1,'',true);
  TOBCLI.SelectDB('',QQ);
  ferme (QQ);
  TOBEDT.AddChampSupValeur ('BPCODECLIENT',TIERS.text);
  TOBEDT.AddChampSupValeur ('BPCLIENTNOM', TOBCLi.GetValue('T_LIBELLE'));
  TOBEDT.AddChampSupValeur ('BPCLIENTADR1', TOBCLi.GetValue('T_ADRESSE1'));
  TOBEDT.AddChampSupValeur ('BPCLIENTADR2', TOBCLi.GetValue('T_ADRESSE2'));
  TOBEDT.AddChampSupValeur ('BPCLIENTADR3', TOBCLi.GetValue('T_ADRESSE3'));
  TOBEDT.AddChampSupValeur ('BPCLIENTCP', TOBCLi.GetValue('T_CODEPOSTAL'));
  TOBEDT.AddChampSupValeur ('BPCLIENTVILLE', TOBCLi.GetValue('T_VILLE'));
  TOBEDT.AddChampSupValeur ('BPLIVNOM', TOBCLi.GetValue('T_LIBELLE'));
  TOBEDT.AddChampSupValeur ('BPLIVADR1', TOBCLi.GetValue('T_ADRESSE1'));
  TOBEDT.AddChampSupValeur ('BPLIVADR2', TOBCLi.GetValue('T_ADRESSE2'));
  TOBEDT.AddChampSupValeur ('BPLIVADR3', TOBCLi.GetValue('T_ADRESSE3'));
  TOBEDT.AddChampSupValeur ('BPLIVCP', TOBCLi.GetValue('T_CODEPOSTAL'));
  TOBEDT.AddChampSupValeur ('BPLIVVILLE', TOBCLi.GetValue('T_VILLE'));
  TOBEDT.AddChampSupValeur ('BPCODEINTERNE', CODEINTERNEP.text);
  for indice := 0 to fTOBparc.detail.count -1 do
  begin
  	TOBP := TOB.Create ('PARCTIERS',TOBEdt,-1);
    TOBP.Dupliquer(fTOBPARC.detail[Indice],true,true);
  end;
  LanceEtatTOB ('E','BP1','Z01',TOBEdit,true,false,false,nil,'','',false);
  TOBEdit.free;
  TOBCli.free;
end;


procedure TOF_BTPARCTIER.BinsMultipleClick(Sender: Tobject);
var stargument : string;
		TOBL : TOB;
    Arow : integer;
    ART : ThcritmaskEdit;
    stFiche,StWhere,CodeArticleP : string;
begin
 	TOBL := GetTOBPARC (GS.row);
  if TOBL <> nil then  fmemoKey.MemoriseClef (TOBL);
  StArgument := ConstitueArg (nil,taCreat);
  //
	ART := THCritMaskEdit.Create (ecran); ART.Visible := false;
  ART.Text := '';
  //
  stFiche := 'BTARTPARC_RECH';
	stWhere := GetTypeArticleParc;
  StWhere := 'GA_CODEARTICLE=;MULTISELECTION;XX_WHERE=AND '+stWhere;
  DispatchRecherche (ART, 1, '',stWhere, stFiche);
  if TheTOB <> nil then
  begin
    AjouteElementsParcs(TheTOB);
    theTOB.free;
  end;
  TheTOB := nil;
  ART.free;
  //
  SelectClick (self);
  Arow := fmemokey.RetrouvePosition (fTOBparc);
  if Arow < 0 then
  begin
  	Arow := 1
  end else inc(Arow);
  positionne (Arow);
end;


function TOF_BTPARCTIER.FindMaxIndice : integer;
var QQ : TQuery;
		Req : string;
    Niv : integer;
begin
	result := 0;
  if valeur(XX_N4.text) > 0 then
  begin
  	niv := 5;
  end else if valeur(XX_N3.text) > 0 then
  begin
  	niv := 5;
  end else if valeur(XX_N2.text) > 0 then
  begin
  	niv := 3;
  end else if valeur(XX_N1.text) > 0 then
  begin
  	niv := 2
  end else
  begin
    niv := 1;
  end;
	req := 'SELECT MAX(BP1_N'+inttoStr(Niv)+') AS LUILA FROM PARCTIERS WHERE '+	ConstitueRequete;
  QQ := OpenSql (Req,true);
  if not QQ.eof then result := QQ.findField('LUILA').AsInteger;
  ferme (QQ);
end;

procedure TOF_BTPARCTIER.AjouteElementsParcs(UneTOBArt: TOB);
var TOBDetails,TOBD,TOBA : TOB;
		indice : integer;
    QQ : TQuery;
    lastIndice : integer;
begin
  TOBA := TOB.Create ('ARTICLE',nil,-1);
  TOBD := TOB.Create ('PARCTIERS',nil,-1);
  //
  for Indice := 0 TO UneTOBArt.detail.count -1 do
  begin
  	QQ := OpenSql('SELECT * FROM ARTICLE '+
  							  'LEFT JOIN ARTICLECOMPPARC ON BCP_ARTICLE=GA_ARTICLE '+
    							'WHERE GA_ARTICLE="'+
    			UneTOBArt.detail[indice].getValue('ARTICLE')+'"',true,1,'',true);
  	TOBA.InitValeurs(true);
    if not QQ.eof then
    begin
    	TOBA.SelectDB('',QQ);
      TOBD.InitValeurs(false);
      lastIndice := FindMaxIndice;
      ConstitueEltParcFromPere (fTOBPere,TOBD,TOBA,1,lastIndice);
      TOBD.InsertDB(nil);
    end;
    ferme (QQ);
  end;
  //
  FreeAndNil(TOBA);
  FreeAndNil(TOBD);
end;

function TOF_BTPARCTIER.ConstitueRequetePere: string;
begin
	result := '';
	if XX_N1.text = '0' then exit;
	result := 'BP1_TIERS="'+TIERS.text+'"';
  //
  if valeur(XX_N5.text) > 0 then
  begin
  	result := result +
    					' AND BP1_N5='+XX_N5.text+
    					' AND BP1_N4='+XX_N4.text+
    					' AND BP1_N3='+ XX_N3.text+
    					' AND BP1_N2='+ XX_N2.text+
    					' AND BP1_N1='+ XX_N1.text;
  end else if valeur(XX_N4.text) > 0 then
  begin
  	result := result + ' AND BP1_N4='+XX_N4.text+
    					' AND BP1_N3='+ XX_N3.text+
    					' AND BP1_N2='+ XX_N2.text+
    					' AND BP1_N1='+ XX_N1.text;
    result := result + ' AND BP1_N5 <> 0';
  end else if valeur(XX_N3.text) > 0 then
  begin
  	result := result + ' AND BP1_N3='+ XX_N3.text+
    					' AND BP1_N2='+ XX_N2.text+
    					' AND BP1_N1='+ XX_N1.text;
    if not TOUTPARC.Checked then result := result + 'AND BP1_N4<>0 AND BP1_N5=0';
  end else if valeur(XX_N2.text) > 0 then
  begin
  	result := result + ' AND BP1_N2='+ XX_N2.text+
    					' AND BP1_N1='+ XX_N1.text;
  end else if valeur(XX_N1.text) > 0 then
  begin
  	result := result + ' AND BP1_N1='+ XX_N1.text;
  end;
end;

procedure TOF_BTPARCTIER.ChargeTOBpere;
var QQ : TQuery;
begin
	if freqPere <> '' then
  begin
    QQ := OpenSql ('SELECT * FROM PARCTIERS WHERE '+freqPere,true,1,'',true);
  	fTOBpere.selectDB ('',QQ);
    Ferme(QQ);
  end else
  begin
  	fTOBPere.putValue('BP1_TIERS',TIERS.text);
  end;
end;

Initialization
  registerclasses ( [ TOF_BTPARCTIER ] ) ;
end.
