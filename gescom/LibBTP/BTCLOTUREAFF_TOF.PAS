{***********UNITE*************************************************
Auteur  ...... : LS
Créé le ...... : 16/03/2009
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCLOTUREAFF ()
Mots clefs ... : TOF;BTCLOTUREAFF
*****************************************************************}
Unit BTCLOTUREAFF_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul, fe_main,
{$else}
     eMul,MainEagl,
{$ENDIF}
		 AglInit,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     ExtCtrls,
     HCtrls,
     Graphics,
     ImgList,
     HEnt1,
     HMsgBox,
     HTB97,
     splash,
     grids,
     types,
     menus,
     UTOF ;

Type
  //
  TOF_BTCLOTUREAFF = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	ImFlags,ImInfo : TImageList;
    AFFAIRE,STATUS,DESCERREUR : integer;
  	MERREURS : TMemo;
    TBSTATUS,TBDESCERR : TToolBarButton97;
    Mnrefus,MnAccept : TmenuItem;
    POPDEMANDE,POPFORCE : TPopupMenu;
    GS : THGrid;
    modeSolo : boolean;
    ListeChamps : string;
    TheFenInfo : TToolWindow97;

    procedure DefinieGrille;
    procedure DefiniAttributCol(Nom: string; Colonne: integer; LaLargeur,Lalignement, Letitre, LeNC: string);
    procedure GSPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    function RecupStatus(TOBL: TOB): integer;
    procedure ShowButtons(Arow: integer; state: boolean);
    procedure MnrefusClick (Sender : Tobject);
    procedure MnAcceptClick (Sender : TObject);
    procedure MnForceClotureClick(Sender: TObject);
    procedure MnPasForceClick(Sender: Tobject);
    procedure TBDESCERRClick (Sender : Tobject);
    procedure definiToolBar;
  end ;

procedure AlimenteErreur(TOBERR: TOB; TheErreur: TMemo);

//FV1 : 11/09/2013 - FS#654 - BAGE : en cloture de chantier, ajouter la saisie de la date de clôture
procedure AddClotureChantierMul(TheAffaire : string; TOBAffaires : TOB; DateCloture : TDateTime);
function ControleChantierSolo(TheAffaire: string;var UReason: integer;  DateCloture : TDateTime) : integer;
function AutoriseClotureChantier (TheAffaire: string; Ureason : integer;DateCloture : TDateTime) : boolean;
procedure LanceClotureChantierMul(XX : TForm;TOBChantiers : TOB; DateCloture : TDateTime);
procedure LanceDeClotureSolo (TheAffaire : string);
procedure LanceTraitementClotureSolo (TheAffaire : string; DateCloture : TDateTime);
function LanceTraitementClotureChantierMul (XX : TForm; TOBCHantiers : TOB; DateCloture : TDateTime) : integer;


Implementation

function LanceTraitementClotureChantierMul (XX : TForm; TOBCHantiers : TOB; DateCloture : TDateTime) : integer;
var splash  : TFsplashScreen;
		Indice : integer;
    TheAffaire : string;
    TOBL : TOB;
begin
	result := 0;
	splash := TFsplashScreen.Create (XX);
  BEGINTRANS;
  TRY
    TRY
      for Indice := 0 to TOBChantiers.detail.count -1 DO
      BEGIN
      	TOBL := TOBchantiers.detail[Indice];
        if (TOBL.getValue ('CHOIX RETOUR')< 0) or (TOBL.getValue ('CHOIX RETOUR')>= 254) then
        BEGIN
        	continue;
        END;
        TheAffaire := TOBL.getValue('AFFAIRE');
        splash.Label1.Caption := 'Clôture du chantier '+ TheAffaire+' en cours...';
        splash.Show;
        splash.Refresh;
        LanceTraitementClotureSolo (TheAffaire, DateCloture);
        inc(result);
      END;
      COMMITTRANS;
    EXCEPT
      ROLLBACK;
      result := -1;
      exit;
    END;
  FINALLY
    splash.Free;
  END;
end;

procedure LanceTraitementClotureSolo (TheAffaire : string; DateCloture : TDateTime);
begin

	ExecuteSQL ('UPDATE PIECE SET GP_VIVANTE="-" WHERE GP_AFFAIRE="'+TheAffaire+'" AND GP_NATUREPIECEG IN ("PBT","CBT","CF","CFR","LFR","BLF","BCE","PBT","LBT","FF","AF","AFS","BFC")');

  ExecuteSql ('UPDATE LIGNE SET GL_VIVANTE="-" WHERE GL_AFFAIRE="'+TheAffaire+'" '+
  						'AND GL_NATUREPIECEG IN ("PBT","CBT","CF","CFR","LFR","BLF","BCE","PBT","LBT","FF","AF","AFS","BFC")');

  ExecuteSql ('UPDATE AFFAIRE SET AFF_ETATAFFAIRE="TER", AFF_DATEFIN="'+ USDATETIME(DateCloture) + '" WHERE AFF_AFFAIRE="'+TheAffaire+ '"');

  // verrouillage des devis non acceptes sur le chantier
	ExecuteSQL ('UPDATE PIECE SET GP_VIVANTE="-" WHERE GP_NATUREPIECEG="DBT" AND GP_NUMERO IN '+
  						'(SELECT GP_NUMERO FROM PIECE WHERE GP_NATUREPIECEG="DBT" AND GP_AFFAIRE="'+TheAffaire+
              '" AND (SELECT A1.AFF_ETATAFFAIRE FROM AFFAIRE A1 WHERE A1.AFF_AFFAIRE=GP_AFFAIREDEVIS)="ENC" '+
              'AND (SELECT A2.AFF_ETATAFFAIRE FROM AFFAIRE A2 WHERE A2.AFF_AFFAIRE=GP_AFFAIRE)="TER")');

	ExecuteSQL ('UPDATE LIGNE SET GL_VIVANTE="-" WHERE GL_NATUREPIECEG="DBT" AND GL_NUMERO IN '+
  						'(SELECT GP_NUMERO FROM PIECE WHERE GP_NATUREPIECEG="DBT" AND GP_AFFAIRE="'+TheAffaire+
              '" AND (SELECT A1.AFF_ETATAFFAIRE FROM AFFAIRE A1 WHERE A1.AFF_AFFAIRE=GP_AFFAIREDEVIS)="ENC" '+
              'AND (SELECT A2.AFF_ETATAFFAIRE FROM AFFAIRE A2 WHERE A2.AFF_AFFAIRE=GP_AFFAIRE)="TER")');

end;

procedure LanceDeClotureSolo (TheAffaire : string);
begin
  ExecuteSql ('UPDATE AFFAIRE SET AFF_ETATAFFAIRE="ACP", AFF_DATEFIN="'+ USDATETIME(idate2099) +'" WHERE AFF_AFFAIRE="'+TheAffaire+ '"');
end;

procedure AddClotureChantierMul(TheAffaire : string; TOBAffaires : TOB; DateCloture : Tdatetime);
var TOBCHA : TOB;
		QQ : TQuery;
    TheLibelle : string;
begin
	QQ := OPenSql ('SELECT AFF_LIBELLE FROM AFFAIRE WHERE AFF_AFFAIRE="'+TheAffaire+'"',True,-1,'',true);
  if not QQ.eof then theLibelle := QQ.findField('AFF_LIBELLE').AsString;
  ferme (QQ);
	TOBCHA := TOB.Create ('UN CHANTIER',TOBAffaires,-1);
  TOBCHA.AddChampSupValeur('AFFAIRE',TheAffaire);
  TOBCHA.AddChampSupValeur('LIBELLE',Thelibelle);
  TOBCHA.AddChampSupValeur('DATECLOTURE',DateToStr(DateCloture));
  TOBCHA.AddChampSupValeur('REASON',0);
  TOBCHA.AddChampSupValeur('DESCERREUR','');
  TOBCHA.AddChampSupValeur('CHOIX RETOUR',-1);
end;

procedure LanceClotureChantierMul(XX : TForm;TOBChantiers : TOB; DateCloture : TDateTime);
var Indice      : integer;
		Ureason     : integer;
		splash      : TFsplashScreen;
    TheAffaire  : string;
    MERREURS    : TMemo;
begin
	MERREURS := Tmemo.create(XX);
  MERREURS.name := 'MESERREURS';
  MERREURS.Parent := XX;
  MERREURS.Visible := false;
	splash := TFsplashScreen.Create (XX);
  //
  for Indice := 0 to TOBChantiers.detail.count -1 DO
  BEGIN
	  TheAffaire := TOBchantiers.detail[Indice].getValue('AFFAIRE');
    splash.Label1.Caption := 'Controle du chantier '+ TheAffaire+' en cours...';
    splash.Show;
    splash.Refresh;
    ControleChantierSolo (TheAffaire,Ureason, DateCloture);
    TOBchantiers.detail[Indice].putvalue('REASON',Ureason);
    AlimenteErreur (TOBchantiers.detail[Indice],MErreurs);
    TOBchantiers.detail[Indice].putvalue('DESCERREUR',MErreurs.text);
  	if Ureason = 0    then TOBchantiers.detail[Indice].putValue('CHOIX RETOUR',0);
    if Ureason >= 254 then TOBchantiers.detail[Indice].putValue('CHOIX RETOUR',-1);
  END;
  splash.Free;
  MErreurs.free;
  TheTOB := TOBChantiers;
  AGLLanceFiche('BTP','BTCLOTUREAFF','','','MULTI');
  TheTOB := nil;

end;

function AutoriseClotureChantier (TheAffaire: string; Ureason : integer; DateCloture : TDateTime) : boolean;
var UneTOB : TOB;
begin
	result := false;
  UneTOB := TOB.Create ('LA LIAISON',nil,-1);
  TRY
    UneTOB.AddChampSupValeur('THE AFFAIRE',TheAffaire);
    UneTOB.AddChampSupValeur('DATECLoture',DateToStr(DateCloture));
    UneTOB.AddChampSupValeur('REASON',UReason);
    UneTOB.AddChampSupValeur('MODE','SOLO');
    UneTOB.AddChampSupValeur('CHOIX RETOUR',-1); // Par defaut annule
    TheTOB := UneTOB;
    AGLLanceFiche('BTP','BTCLOTUREAFF','','','SOLO');
    if UneTOB.getValue('CHOIX RETOUR')=0 then result := true;
  FINALLY
  	UneTOB.free;
  END;
end;

function ControleChantierSolo(TheAffaire: string;var UReason: Integer; DateCloture : TDateTime): integer;
var REq : string;
		QQ : TQuery;
begin
	result := 0;
  Ureason := 0;
  //
  req := 'SELECT SUM(BCO_QUANTITE) AS RECEPT ,SUM(BCO_QTEVENTE) AS LIVRED '+  // ^^
         'FROM CONSOMMATIONS '+
         'WHERE BCO_AFFAIRE="'+TheAffaire+ '" AND BCO_TRAITEVENTE<>"X" '+
				 'AND BCO_TRANSFORME<>"X" '+
				 'AND BCO_NATUREPIECEG IN ("BLF","LFR","FF")';
  QQ := OpenSQL  (Req,true,-1,'',true);
  if not QQ.eof then
  begin
    if (QQ.findField('RECEPT').AsFloat - QQ.FindField('LIVRED').AsFloat > 0 ) then
    begin
    	Ureason := 255;
      result := 2;
    end;
  end;
  ferme (QQ);

  // Traitement de decisionnel d'achat
  if result < 2 then
  begin
    req := 'SELECT BAD_QUANTITEINIT,BAE_VALIDE '+  // ^^
           'FROM DECISIONACHLIG '+
           'LEFT JOIN DECISIONACH ON BAE_NUMERO=BAD_NUMERO '+
           'WHERE BAD_AFFAIRE="'+TheAffaire+ '" AND BAE_VALIDE="-"';
    QQ := OpenSQL  (Req,true,-1,'',true);
    if not QQ.eof then
    begin
      Ureason := 254;
      result := 2;
    end;
    ferme (QQ);
  end;
  //
  if result < 2 then
  begin
    //
    //
  	req := 'SELECT GL_NATUREPIECEG FROM LIGNE WHERE GL_NATUREPIECEG IN ("CF","CFR") '+
    			 'AND (GL_QTERESTE <> 0 OR GL_MTRESTE<>0) AND GL_AFFAIRE="'+TheAffaire+'" AND GL_SOLDERELIQUAT="-" and GL_VIVANTE="X"';
    if ExisteSQL(req) then
    begin
    	Ureason := 1;
      result := 1;
    end;
    //
    REq := 'SELECT GL_NATUREPIECEG FROM LIGNE WHERE GL_NATUREPIECEG="PBT" '+
    			 'AND (GL_QTERESTE <> 0 OR GL_MTRESTE<>0) AND GL_AFFAIRE="'+TheAffaire+'" AND GL_SOLDERELIQUAT="-" and GL_VIVANTE="X" AND '+
           '(SELECT BNP_TYPERESSOURCE FROM ARTICLE LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES '+
           'WHERE GA_ARTICLE=GL_ARTICLE) NOT IN ("SAL","MAT")' ;
    if ExisteSQL(req) then
    begin
    	Ureason := Ureason + 2;
      result := 1;
    end;
    //
    REq := 'SELECT GL_NATUREPIECEG FROM LIGNE WHERE GL_NATUREPIECEG="CBT" '+
    			 'AND (GL_QTERESTE <> 0 OR GL_MTRESTE<>0) AND GL_AFFAIRE="'+TheAffaire+'" AND GL_SOLDERELIQUAT="-" and GL_VIVANTE="X"';
    if ExisteSQL(Req) then
    begin
    	Ureason := ureason + 4;
      result := 1;
    end;
    //
    //FV1 : 11/09/2013 - FS#654 - BAGE : en cloture de chantier, ajouter la saisie de la date de clôture
    req := 'SELECT BCO_DATEMOUV FROM CONSOMMATIONS ' +
           'WHERE BCO_AFFAIRE="' + TheAffaire +
           '" AND BCO_DATEMOUV>"'+ USDATETIME(DateCloture) + '"';
    if ExisteSQL(req) then
    begin
    	Ureason := ureason + 8;
      result := 1;
    end;
    //
  end;
  //
end;

procedure TOF_BTCLOTUREAFF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREAFF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREAFF.OnUpdate ;
var Indice : integer;
		BCallDayPass : Boolean;
    TOBPass : TOB;
begin
  Inherited ;
  if ModeSolo then
  begin
		LaTob.putvalue('CHOIX RETOUR',0);
  end else
  begin
  	BCallDayPass := false;
    for Indice := 0 to LatoB.detail.count -1 do
    begin
      if (LaTOB.detail[Indice].getValue('REASON')>=254) and (LaTOB.detail[Indice].getValue('CHOIX RETOUR')=1) then
      begin
      	BCallDayPass := true;
        break;
      end;
    end;
    if BCallDayPass then
    begin
    	TOBPass := TOB.create ('UNE TOB',nil,-1);
      TOBPASS.AddChampSupValeur ('DAYPASS','');
      TOBPASS.AddChampSupValeur ('RESULTAT','');
      TheTOB := TOBPAss;
      AGLLAnceFiche('BTP','BTCONFIRMPASS','','','');
      TheTOB := nil;
      if TOBPass.getValue('RESULTAT')<>'OK' then LaTob.putvalue('OK','-') else LaTob.putvalue('OK','X');
      TOBPass.free;
    end else
    begin
			LaTob.putvalue('OK','X');
    end;
  end;
end ;

procedure TOF_BTCLOTUREAFF.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREAFF.OnArgument (S : String ) ;
var UneImage : THimage;
		cAncel : boolean;
begin
  Inherited ;
  cancel := false;
  if pos('SOLO',S)  >0 then ModeSolo := true else ModeSolo := false;
  MERREURS := TMemo(GetControl('MERREURS'));
  GS := THgrid(GetControl('GS'));

  Ecran.Caption := 'Clôture des Affaire à la date du : ' + LaTob.detail[0].GetValue('DATECLOTURE');

  if modeSolo then
  begin
  	GS.visible := false;
    AlimenteErreur (LaTOB,MErreurs);
    Merreurs.ReadOnly := true;
  	if LaTOB.getValue ('REASON')>=254 then TToolbarButton97 (GetControl('BVALIDER')).visible := false;
    TBStatus := nil;
    TBDESCERR := nil;
  	ImFlags := nil;
    ImInfo := nil;
  end else
  begin
    POPDEMANDE := TPopupMenu(getControl('POPDEMANDE'));
    POPFORCE := TPopupMenu(getControl('POPFORCE'));
  	TBStatus := TToolBarButton97(GetCOntrol('TBSTATUS'));
    TBDESCERR := TToolBarButton97(GetCOntrol('TBDESCERR'));
    TBDESCERR.onclick := TBDESCERRClick;
    Mnrefus := TmenuItem(getCOntrol('Mnrefus'));
    MnAccept := TmenuItem(getCOntrol('MnAccept'));
    Mnrefus.onclick := MnrefusClick;
    MnAccept.OnClick := MnAcceptClick;
    TmenuItem(GetCOntrol('MnForceCloture')).OnClick := MnForceClotureClick;
    TmenuItem(GetCOntrol('MnpasForce')).OnClick := MnPasForceClick;
    //
  	ImFlags := TImageList.create(GS);
    ImFlags.Height := 24;
    ImFlags.Width := 24;
    ImInfo := TImageList.create(GS);
    ImInfo.Height := 24;
    ImInfo.Width := 24;
    //
    UneImage := THImage(GetControl('IERROR'));
    ImFlags.add(TBitMap(UneImage.Picture.Bitmap ),nil);
    UneImage := THImage(GetControl('IOK'));
    ImFlags.add(TBitMap(UneImage.Picture.Bitmap ),nil);
    UneImage := THImage(GetControl('IDEMANDE'));
    ImFlags.add(TBitMap(UneImage.Picture.Bitmap ),nil);
    UneImage := THImage(GetControl('IPASSTOOK'));
    ImFlags.add(TBitMap(UneImage.Picture.Bitmap ),nil);
    UneImage := THImage(GetControl('IPASTOANNUL'));
    ImFlags.add(TBitMap(UneImage.Picture.Bitmap ),nil);

    UneImage := THImage(GetControl('IMESSAGE'));
    ImInfo.add(TBitMap(UneImage.Picture.Bitmap ),nil);
    //
    TBSTATUS.Images := ImFlags;
    TBDESCERR.images := ImInfo;

    GS.PostDrawCell := GSPostDrawCell;
    GS.OnRowEnter := GSRowEnter;
    GS.OnRowExit := GSRowExit;
    DefinieGrille;
    //
    definiToolBar;

    //
    GS.row := 1;
    GS.SetFocus;
    GSRowEnter(self,GS.row,cancel,false)
  end;
end ;

procedure TOF_BTCLOTUREAFF.OnClose ;
begin
	if not modeSolo then
  begin
    ImFlags.free;
    ImInfo.free;
    TheFenInfo.free;
  end;
  Inherited ;
end ;

procedure TOF_BTCLOTUREAFF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREAFF.OnCancel () ;
begin
  Inherited ;
end ;

procedure AlimenteErreur (TOBERR : TOB; TheErreur : TMemo);
begin
	TheErreur.Clear;
  if TOBERR.getValue('REASON') = 255 then
  begin
		TheErreur.Lines.add('              ATTENTION');
		TheErreur.Lines.add('');
		TheErreur.Lines.add('Il existe des articles réceptionnés non livrés sur le chantier');
		TheErreur.Lines.add	('');
		TheErreur.Lines.add('Clôture Impossible');
  end else
  begin
    if TOBERR.getValue('REASON') = 254 then
    begin
      TheErreur.Lines.add('              ATTENTION');
      TheErreur.Lines.add('');
      TheErreur.Lines.add('Il existe des décisionnels d''achat non traités sur le chantier');
      TheErreur.Lines.add	('');
      TheErreur.Lines.add('Clôture Impossible');
    end else
    begin
      if TOBERR.getValue('REASON') = 0 then
      begin
        TheErreur.Lines.add('Aucun problème détecté');
      end else
      begin
        TheErreur.Lines.add('              ATTENTION');
        TheErreur.Lines.add('');
        if TOBERR.getValue('REASON') and 1 = 1 then 	TheErreur.Lines.add('Il existe des commandes fournisseurs non réceptionnées');
        if TOBERR.getValue('REASON') and 2 = 2 then 	TheErreur.Lines.add( 'Il existe des articles non traités sur la prévision de chantier');
        if TOBERR.getValue('REASON') and 4 = 4 then 	TheErreur.Lines.add('Il existe des besoins de chantier non traités');
        if TOBERR.getValue('REASON') and 8 = 8 then 	TheErreur.Lines.add('Il existe des consommations après la date de clôture');
      end;
    end;
  end;
end;

procedure TOF_BTCLOTUREAFF.DefinieGrille;
var
    lesElements,Lelement,LaLargeur,Lalignement,LeTitre,LeNC : string;
    colonne : integer;
    flargeur,Falignement,FTitre,NC : string;
    TOBL : TOB;
    Indice : integer;
begin

  ListeChamps := 'INDICE;AFFAIRE;LIBELLE;REASON;DESCERREUR;';
  flargeur := '2;10;30;6;6;';
  Falignement := 'G.0O X--;G.0  ---;G.0  ---;C.0  ---;G.0  ---;';
  FTitre := 'N°;Chantier;Désignation;Status;Raison;';
  NC := '0;1;0;0;0;'; //definition des zones saisissable ou non

  GS.ColCount := 5;
  lesElements := ListeChamps;
  //
  Lelement := ReadtokenSt(lesElements);
  LaLargeur := ReadtokenSt(fLargeur);
  Lalignement := ReadtokenSt(Falignement);
  LeTItre := readTokenSt(Ftitre);
  LeNC := readTokenSt(NC);
  Colonne := 0;
  Repeat
    if Lelement = '' then break;
    if Lelement = 'AFFAIRE' then
    begin
      AFFAIRE := Colonne;
      DefiniAttributCol (Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end else
    begin
    	if Lelement = 'REASON' then STATUS := colonne;
    	if Lelement = 'DESCERREUR' then DESCERREUR := colonne;
      DefiniAttributCol (Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end;
    //
    Lelement := ReadtokenSt(lesElements);
    LaLargeur := ReadtokenSt(fLargeur);
    Lalignement := ReadtokenSt(Falignement);
    LeTItre := readTokenSt(Ftitre);
    LeNC := readTokenSt(NC);
    inc(colonne);
    //
  until lelement = '';
  GS.RowCount := LaTOB.detail.count+1;
  for Indice := 0 to LaTOB.detail.count -1 do
  begin
  	TOBL :=  LaTOB.detail[Indice];
  	GS.RowHeights [Indice+1] := 28;
  	GS.cells[0,Indice+1] := InttOStr(Indice+1);
    GS.cells[1,Indice+1] := TOBL.getValue('AFFAIRE');
    GS.cells[2,Indice+1] := TOBL.getValue('LIBELLE');
    GS.cells[STATUS,Indice+1] := '';
    GS.cells[DESCERREUR,Indice+1] := '';
  end;
end;


procedure TOF_BTCLOTUREAFF.DefiniAttributCol(Nom: string; Colonne: integer; LaLargeur,Lalignement, Letitre, LeNC: string);
var FF: string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep : boolean;
    dec: integer;

begin
  TransAlign(Lalignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
  GS.cells[Colonne,0] := leTitre;
  // Alignement dans la colonne
  if copy(LAlignement,1,1)='G' then GS.ColAligns[Colonne] := taLeftJustify
  else if copy(LAlignement,1,1)='D' then GS.ColAligns[Colonne] := taRightJustify
  else if copy(LAlignement,1,1)='C' then GS.ColAligns[Colonne] := taCenter;
  GS.ColWidths[Colonne] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
  if OkLib then
  begin
    GS.ColFormats[Colonne] := 'CB=' + Get_Join(Nom);
  end else if (Dec<>0) or (Sep) then
  begin
  	GS.ColFormats[Colonne] := FF;
  end;
  if LeNc = '0' then GS.ColEditables [Colonne]:= false;
end;

function TOF_BTCLOTUREAFF.RecupStatus(TOBL : TOB) : integer;
begin
	result := 0;
	if (TOBL.getValue('REASON')>=254)      and (TOBL.getValue('CHOIX RETOUR')=-1) then
    result := 0
  else if (TOBL.getValue('REASON')>=254) and (TOBL.getValue('CHOIX RETOUR')=1)  then
    result := 3
  else if (TOBL.GetValue('REASON')=0)                                           then
    result := 1
  else if (TOBL.getValue('REASON')>0)    and (TOBL.getValue('CHOIX RETOUR')=0)  then
    result := 2
  else if (TOBL.getValue('REASON')>0)    and (TOBL.getValue('CHOIX RETOUR')=1)  then
    result := 3
  else if (TOBL.getValue('REASON')>0)    and (TOBL.getValue('CHOIX RETOUR')=-1) then
    result := 4;

end;

procedure TOF_BTCLOTUREAFF.GSPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var ARect: TRect;
		Numgraph : integer;
    TOBL : TOB;
begin
  ARect := GS.CellRect(ACol, ARow);
  GS.Canvas.Pen.Style := psSolid;
  GS.Canvas.Pen.Color := clgray;
  GS.Canvas.Brush.Style := BsSolid;
  //
  if (Acol = STATUS) and (Arow >= GS.fixedRows) then
  begin
  	TOBL := LaTOB.detail[Arow-1];
    if TOBL = nil then Exit;
    Canvas.FillRect(ARect);
    NumGraph := RecupStatus(TOBL);
    if NumGraph >= 0 then
    begin
      ImFlags.DrawingStyle := dsTransparent;
      ImFlags.Draw(Canvas, ((Arect.left+Arect.Right)- 24) div 2, ARect.top+2, NumGraph);
    end;
  end;
  //
  if (Acol = DESCERREUR) and (Arow >= GS.fixedRows) then
  begin
  	TOBL := LaTOB.detail[Arow-1];
    if TOBL = nil then Exit;
    Canvas.FillRect(ARect);
    ImInfo.DrawingStyle := dsTransparent;
    ImInfo.Draw(Canvas, ((Arect.left+Arect.Right)- 24) div 2, ARect.top+2, 0);
  end;
end;

procedure TOF_BTCLOTUREAFF.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var TOBL : TOB;
begin
  if (GS.Row < GS.fixedRows) then exit;
 	TOBL := LaTOB.detail[GS.row-1];
  MERREURS.text := TOBL.Getvalue('DESCERREUR');
	ShowButtons (GS.row,true);
end;

procedure TOF_BTCLOTUREAFF.GSRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
	ShowButtons (Ou,false);
end;

procedure TOF_BTCLOTUREAFF.ShowButtons (Arow : integer; state : boolean);
var arect: Trect;
		Numgraph : integer;
    TOBL : TOB;
begin
  if (Arow < GS.fixedRows) then exit;
	if state then
  begin
  	TOBL := LaTOB.detail[Arow-1];
    if (TOBL.GetValue('REASON') > 0)  then
    begin
			ARect := GS.CellRect(STATUS, Arow);
      NumGraph := RecupStatus(TOBL);
      TBSTATUS.ImageIndex := NumGraph;
      TBSTATUS.Opaque := false;
      with TBSTATUS do
      begin
        Top := Arect.top - GS.GridLineWidth;
        Left := Arect.Left;
        Width := Arect.Right - Arect.Left;
        Height := Arect.Bottom - Arect.Top;
        Parent := GS;
        Visible := true;
        if TOBL.GetValue('REASON') < 254 then
        begin
          DropdownArrow := true;
          DropdownMenu := POPDEMANDE;
        end else if TOBL.GetValue('REASON') >= 254 then
        begin
          DropdownArrow := true;
          DropdownMenu := POPFORCE;
        end else
        begin
          DropdownArrow := false;
          DropdownMenu := nil;
        end;
      end;
    end;
    if (TOBL.GetValue('REASON') > 0) then
    begin
    	ARect := GS.CellRect(DESCERREUR, Arow);
    	TBDESCERR.ImageIndex := 0;
    	TBDESCERR.Opaque := false;
      with TBDESCERR do
      begin
        Top := Arect.top - GS.GridLineWidth;
        Left := Arect.Left;
        Width := Arect.Right - Arect.Left;
        Height := Arect.Bottom - Arect.Top;
        Parent := GS;
        Visible := true;
      end;
    end;
  end else
  begin
    TBSTATUS.visible := false;
    TBDESCERR.visible := false;
  end;
end;

procedure TOF_BTCLOTUREAFF.MnAcceptClick(Sender: TObject);
var TOBL : TOB;
begin
	TOBL := LaTOB.detail[GS.row-1];
  TOBL.putValue('CHOIX RETOUR',1);
  GS.InvalidateRow(GS.row);
  ShowButtons (GS.row,true);
end;

procedure TOF_BTCLOTUREAFF.MnForceClotureClick(Sender: TObject);
var TOBL : TOB;
begin
	// demande passe du jour
  //
	TOBL := LaTOB.detail[GS.row-1];
  TOBL.putValue('CHOIX RETOUR',1);
  GS.InvalidateRow(GS.row);
  ShowButtons (GS.row,true);
end;

procedure TOF_BTCLOTUREAFF.MnrefusClick(Sender: Tobject);
var TOBL : TOB;
begin
	TOBL := LaTOB.detail[GS.row-1];
  TOBL.putValue('CHOIX RETOUR',-1);
  GS.InvalidateRow(GS.row);
  ShowButtons (GS.row,true);
end;

procedure TOF_BTCLOTUREAFF.MnPasForceClick(Sender: Tobject);
var TOBL : TOB;
begin
	TOBL := LaTOB.detail[GS.row-1];
  TOBL.putValue('CHOIX RETOUR',-1);
  GS.InvalidateRow(GS.row);
  ShowButtons (GS.row,true);
end;

procedure TOF_BTCLOTUREAFF.TBDESCERRClick(Sender: Tobject);
var TOBL : TOB;
begin
	TOBL := LaTOB.detail[GS.row-1];
  MERREURS.text := TOBL.Getvalue('DESCERREUR');
  TheFenInfo.Visible := true;
end;

procedure TOF_BTCLOTUREAFF.definiToolBar;
var ARect : TRect;
begin
TheFenInfo := TToolWindow97.create (ecran);
TheFenInfo.parent := ecran;
TheFenInfo.caption := 'Message d''information';
TheFenInfo.BorderStyle := bsSingle;
TheFenInfo.clientAreaHeight :=MERREURS.Height ;
TheFenInfo.clientAreaWidth := MERREURS.Width ;
TheFenInfo.ClientHeight := TheFenInfo.clientAreaHeight;
TheFenInfo.clientWidth := TheFenInfo.clientAreaWidth;
TheFenInfo.DragHandleStyle := dhDouble;
TheFenInfo.HideWhenInactive := True;
TheFenInfo.fullsize := false;
TheFenInfo.Resizable := true;
TheFenInfo.Height := 200;
TheFenInfo.Width := 320;
Arect := Ecran.BoundsRect;
TheFenInfo.top := Arect.top ;
TheFenInfo.left := Arect.left;
MERREURS.Parent := TheFenInfo;
TheFenInfo.visible := false;
end;


Initialization
  registerclasses ( [ TOF_BTCLOTUREAFF ] ) ;
end.
