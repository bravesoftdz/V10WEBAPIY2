unit UtofEditionBI;

{*****************************************************************
Source  utilisé pour les fiches : RTSUSP_MUL_MAILIN
                                  RTPROS_MUL_MAILIN
                                  RTPROS_MAILIN_FIC
                                  RTPRO_MAILIN_CONT
*****************************************************************}

interface

{
	function ConvertDocFile(
		FileSrce, FileDest: String;
		Initialize, Finalize: TNotifyEvent;
		Funct: TOnFuncEvent;
		EGetVar: TGetVarEvent;
		ESetVar: TSetVarEvent;
		GetList: TOnGetListEvent ): Integer;
}

uses {$IFDEF VER150} variants,{$ENDIF}  Controls,
		  Classes,
      forms,
      sysutils,
      ComCtrls,
      HCtrls,
      HMsgBox,
      UTOF,
      HQry,
      HEnt1,
      StdCtrls,
{$IFDEF EAGLCLIENT}
			UtileAGL,
      eMul,
      MaineAGL,
{$ELSE}
      mul,
      Fe_Main,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      EdtREtat,
      EdtRDoc,
{$ENDIF}
			Doc_Parser,
      RT_Parser,
      Hstatus,
      HDB,
      M3FP,
      HTB97,
      UTOB,
      UtilSelection,
      Paramsoc,
      UtilPGI,
      UtilRT,
      UtilWord,
      AGLUtilOLE,
      UtilGC,
      UtilArticle,
      UtilsEtat,
      Ent1,
      AffaireUtil;

Type
    TOF_EditionBI = Class (TOF)
    Private
        TobProsp      : TOB ;

        BMailingDoc   : TToolbarButton97;
        BOuvrir       : TToolBarButton97;
        BParamEtat    : TToolBarButton97;

        //
        ChkDuplicata  : TCheckBox;
        ChkApercu     : THCheckBox;
        ChkReduire    : THCheckBox;
        ChkCouleur    : THCheckBox;
        ChkFichierPDF : TCheckBox;
        ChkEnvoiMail  : TCheckBox;
        //
        Aff       		: THEdit;
	      Aff0      		: THEdit;
  	    Aff1      		: THEdit;
    	  Aff2      		: THEdit;
	      Aff3      		: THEdit;
  	    Aff4      		: THEdit;
        Tiers					: THEdit;

        statut				: String;

        stProduitpgi  : string;
        StApplication : String;

        FETAT         : THValComboBox;
        TEtat         : ThLabel;
        //
        Idef          : Integer;
        //
        OptionEdition : TOptionEdition;
        TheType       : String;
        TheNature     : String;
        TheTitre      : String;
        TheModele     : String;
        fRealise 			: boolean;
        Grille				: THDBGrid;

        procedure ChangeDocument;
        procedure ChangeMaquette;
        Procedure ChargeEdition(TobAppel : Tob);
        procedure ChargeEditionAll(TobAppel : tob);
        procedure ControleChamp(Champ : String;Valeur : String);
        procedure LanceEdition;
        procedure Publipostage;

        Function MailingValidDoc : Boolean;
        Function NomCorrect ( NomPDF : String ) : Boolean ;

        procedure BOuvrir_Click(Sender: Tobject);
     	  procedure DoubleClick(sender: TObject);
    	  procedure MailingDoc(sender: TObject);
    procedure ChargeEtat;
    procedure OnClickApercu(Sender: Tobject);
    procedure OnChangeEtat(Sender: Tobject);
    procedure ParamEtat(Sender: TOBJect);
    procedure OnClickDuplicata(Sender: Tobject);
    procedure OnClickCouleur(Sender: Tobject);
    procedure OnClickReduire(Sender: Tobject);
    procedure OnClickFichierPDF(Sender: Tobject);

    Public
        procedure OnArgument (Arguments : String ); override ;
        procedure OnLoad ; override ;
        procedure EditionBI(Maquette , DocGenere : string);
        procedure OnGetVar( Sender: TObject; VarName: String; VarIndx: Integer;var Value: variant) ;
     END;

procedure BTPEditionBI(parms:array of variant; nb: integer ) ;
Function BTPMailingValidDoc(parms:array of variant; nb: integer ) : variant ;

implementation

uses AGLInitRT;

procedure BTPEditionBI(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin

F:=TForm(Longint(Parms[0])) ;

if (F is TFmul) then
   TOTOF:=TFMul(F).LaTOF
else
   exit;

if (TOTOF is TOF_EditionBI) then
   TOF_EditionBI(TOTOF).EditionBI(string(Parms[1]),string(Parms[2]))
else
   exit;

end;

Function BTPMailingValidDoc( parms: array of variant; nb: integer ) : variant;
var  F : TForm ;
     TOTOF  : TOF;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFmul) then
     TOTOF:=TFMul(F).LaTOF
  else
	  exit;
   result:=TOF_EditionBI(TOTOF).MailingValidDoc;
end;

procedure TOF_EditionBI.OnArgument (Arguments : String );
var	F 						: TForm;
    LaVersionWord : integer;
    NomForme			: String;
    Critere				: String;
		Valeur 				: String;
	  Champ   			: String;
  	X      				: integer;
    CC            : THValComboBox;
begin
inherited ;
  fRealise := false;
	F := TForm (Ecran);
  StApplication := '';
  //Récupération valeur de argument
	Critere:=(Trim(ReadTokenSt(Arguments)));

  while (Critere <> '') do
      begin
      if Critere <> '' then
         begin
         X := pos (':', Critere) ;
         if x = 0 then
            X := pos ('=', Critere) ;
         if x <> 0 then
            begin
            Champ := copy (Critere, 1, X - 1) ;
            Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
        	  ControleChamp(champ, valeur);
				    end
         end;
         Critere := (Trim(ReadTokenSt(Arguments)));
      end;


  // gestion Etablissement (BTP)
	CC:=THValComboBox(GetControl('BTBETABLISSEMENT')) ;
	if CC<>Nil then PositionneEtabUser(CC) ;

  // gestion Domaine (BTP)
	CC:=THValComboBox(GetControl('AFF_DOMAINE')) ;
	if CC<>Nil then PositionneDomaineUser(CC) ;


  SetControlVisible('PCritere',True);
  SetControlVisible('PComplement',FALSE);
  SetControlVisible('PAvance',FALSE);
  SetControlVisible('PSQL',FALSE);
  SetFocusControl('AVEC');

  if not GetParamSoc('SO_EDITWORD') Then
     Begin
     SetControlVisible('PDocument',FALSE);
     SetControlVisible('PEdition',TRUE);
     SetControlVisible('PRESSOURCES',FALSE);
     SetControlVisible('BCREATDOC',FALSE);
     SetControlVisible('BMAILINGDOC',FALSE);
     End
  ELSE
     Begin
     SetControlVisible('PEdition',FALSE);
     SetControlVisible('PDocument',True);
     SetControlVisible('BCREATDOC',True);
     SetControlVisible('BMAILINGDOC',True);
     SetControlText('MAQUETTE',GetParamsoc('SO_BTDIRMAQUETTE') + '\' + GetParamsocSecur('SO_BINOMMAQUETTE', '*.doc'));
     SetControlText('DOCGENERE',GetParamsoc('SO_BTDIRSORTIE') + '\' + GetParamsocSecur('SO_BINOMMAQUETTE', '*.doc'));
     end;

  //Déclarations et procédures des zones ecran
  Grille := THDBGrid(GetControl('FListe'));
  Grille.OnDblClick := DoubleClick;

	BMailingDoc := TToolbarButton97(ecran.FindComponent('BMAILINGDOC'));
	BMailingDoc.onclick := MailingDoc;

	BOuvrir := TToolbarButton97(ecran.FindComponent('BOUVRIR'));
	BOuvrir.onclick := BOuvrir_Click;


  //FV1 : 31/07/2013 - FS#238 - VIVIANE : donner l'accès au modèle de bon d'intervention depuis le pgm
  if Assigned(GetControl('fEtat'))  then FEtat := ThValComboBox(ecran.FindComponent('FEtat'));
  if assigned(FETAT)                then
  begin
  	FETAT.OnChange := OnChangeEtat;
  end;
  if assigned(Getcontrol('TEtat'))  then TEtat := ThLabel(ecran.FindComponent('TEtat'));
  //
  if Assigned(GetControl('BParamEtat'))  then BParamEtat:= TToolbarButton97(ecran.FindComponent('BParamEtat'));
  if Assigned(BParamEtat)   then BParamEtat.OnClick     := ParamEtat;
  //
  if Assigned(GetControl('fApercu'))   then ChkApercu   := THCheckBox(Ecran.FindComponent('fApercu'));
  if assigned(ChkApercu)               then
  begin
  	ChkApercu.Checked     := True;
  	ChkApercu.OnClick := OnClickApercu;
  end;
  //
  if Assigned(GetControl('BDUPLICATA'))then ChkDuplicata:= TCheckBox(Ecran.FindComponent('BDUPLICATA'));
  if assigned(ChkDuplicata)            then
  begin
  	ChkDuplicata.Checked  := False;
  	ChkDuplicata.OnClick := OnClickDuplicata;
  end;
  //
  if Assigned(GetControl('fCouleur'))  then ChkCouleur  := THCheckBox(Ecran.FindComponent('fCouleur'));
  if assigned(ChkCouleur)              then
  begin
  	ChkCouleur.OnClick := OnClickCouleur;
  	ChkCouleur.Checked    := False;
  end;
  //
  if Assigned(GetControl('fReduire'))  then ChkReduire:= THCheckBox(Ecran.FindComponent('fReduire'));
  if assigned(ChkReduire)              then
  begin
  	ChkReduire.Checked    := False;
  	ChkReduire.OnClick := OnClickReduire;
  end;
  //
  if Assigned(GetControl('fApercu'))   then ChkApercu   := THCheckBox(Ecran.FindComponent('fApercu'));
  if assigned(ChkApercu)               then ChkApercu.OnClick := OnClickApercu;
  //
  if Assigned(GetControl('BFichierPDF')) then ChkFichierPDF  := THCheckBox(Ecran.FindComponent('BFichierPDF'));
  if assigned(ChkFichierPDF)             then
  begin
  	ChkFichierPDF.OnClick := OnClickFichierPDF;
  	ChkFichierPDF.Checked := False;
  end;
  if Fetat <> nil then
  begin
    FEtat.Items.Clear;
    FEtat.Values.Clear;

    //chargement de l'état en fonction du type de paiement : si direct Attestation, si regroupement eclatement
    ChargeEtat;

    if fetat.Items.Count > 1 then
    begin
      Fetat.Visible := true;
      Tetat.Visible := true;
    end;
  end;

end;

//appel de l'écran de modification/création d'état
procedure TOF_EditionBI.ParamEtat(Sender : TOBJect);
begin

  OptionEdition.Appel_Generateur;

end;

procedure TOF_EditionBI.OnClickApercu(Sender : Tobject);
begin
  OptionEdition.Apercu  := ChkApercu.checked;
end;

procedure TOF_EditionBI.OnClickDuplicata(Sender : Tobject);
begin
  OptionEdition.Duplicata := Chkduplicata.checked;
end;

procedure TOF_EditionBI.OnClickCouleur(Sender : Tobject);
begin
  OptionEdition.Couleur := ChkCouleur.checked;
end;

procedure TOF_EditionBI.OnClickFichierPDF(Sender : Tobject);
begin
  SetControlVisible('TNOMPDF', ChkFichierPDF.checked);
  SetControlVisible('NOMPDF',  ChkFichierPDF.checked);
end;

procedure TOF_EditionBI.OnClickReduire(Sender : Tobject);
begin
  OptionEdition.DeuxPages := ChkReduire.checked;
end;

procedure TOF_EditionBI.ControleChamp(Champ : String;Valeur : String);
var affaire : string;
begin

  if champ = 'STATUT' then
  begin
    statut := Valeur;
    if Statut = 'APP'      then SetControlText('TAFF_AFFAIRE', 'Code Appel')
    else if Statut = 'INT' then SetControlText('TAFF_AFFAIRE', 'Code Contrat')
    else if Statut = 'AFF' then SetControlText('TAFF_AFFAIRE', 'Code Chantier')
    else if Statut = 'GRP' then SetControlText('TAFF_AFFAIRE', 'Code Affaire')
    else if Statut = 'PRO' then SetControlText('TAFF_AFFAIRE', 'Appel d''Offre')
    else
      SetControlText('TAFF_AFFAIRE', 'Code Affaire');
  end;
  
  if champ= 'ETAT' then
  begin
    if Valeur = 'AFF' then
    begin
      StApplication := 'AFF_ETATAFFAIRE IN ("AFF","ACA")';
    end else if Valeur = 'REA' then
    begin
      fRealise := True;
      StApplication := 'AFF_ETATAFFAIRE NOT IN ("ECO","ECR","ANN","AFF","ACA","ACD")';
    end;

    SetActiveTabSheet('PRESSOURCE');
  end;

	Aff :=THEdit(GetControl('AFF_AFFAIRE'));

	Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
	Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
	Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
	Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
	Aff4:=THEdit(GetControl('AFF_AVENANT'));

	Tiers:=THEdit(GetControl('AFF_TIERS'));

	if (Aff1<>nil) then
  	 BEGIN
   	 if Aff<>Nil then
     	  Affaire:=Aff.Text
   	 else
     		Affaire:='' ;
    //
	   if statut = 'APP' then
  	 		aff0.text := 'W'
   	 else if Statut = 'INT' then
	  	  aff0.text :='I'
	   else if Statut = 'PRO' then
		    aff0.text :='A'
   	 Else
     	  aff0.text :='A';

   	 ChargeCleAffaire(Aff0,Aff1,Aff2,Aff3,Aff4,Nil,TaCreat,Affaire,false);

     END;

end;

procedure TOF_EditionBI.OnLoad;
begin
	inherited ;
  SetControlText('XX_WHERE',StApplication) ;
end;

procedure TOF_EditionBI.EditionBI(Maquette , DocGenere : string);
var  F    : TFMul;
{$IFDEF EAGLCLIENT}
     L    : THGrid;
{$ELSE}
     Ext  : String;
     L    : THDBGrid;
{$ENDIF}
     i            : integer;
     codeprospect : String;
     stClauseWhere: String;
     StSelect     : String;
     StSelect2    : string;
     QQ           : TQuery ;
     Pages        :TPageControl;
begin

  F:=TFMul(Ecran);

  if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné','');
    exit;
  end;

  if PGIAsk('Confirmez-vous le traitement ?','')<>mrYes then exit ;

L:= F.FListe;
Pages:=F.Pages;

stSelect := 'SELECT * FROM (AFFAIRE LEFT JOIN TIERS ON AFF_TIERS = T_TIERS)';
stSelect := stSelect + ' LEFT JOIN RESSOURCE ON AFF_RESPONSABLE = ARS_RESSOURCE';
stSelect := stSelect + ' LEFT JOIN ADRESSES ON AFF_AFFAIRE = ADR_REFCODE ';
stSelect := stSelect + ' WHERE AFF_AFFAIRE0 = "W"';

TobProsp:=Tob.create('_Les Prospects',nil,-1);

if L.AllSelected then
   begin
   stClauseWhere:=RecupWhereCritere(Pages);
   StSelect := StSelect + stClauseWhere;
   QQ:=OpenSQL(StSelect,TRUE) ;
   if Not QQ.EOF then TobProsp.LoadDetailDB('','','',QQ,False,True);
   Ferme(QQ) ;
   L.AllSelected:=False;
   end else
   begin
   InitMove(L.NbSelected,'');
   for i:=0 to L.NbSelected-1 do
      begin
      MoveCur(False);
      L.GotoLeBookmark(i);
      Tob.create('',TobProsp,-1);
      codeprospect:= F.Q.FindField('T_TIERS').asstring ;
      StSelect2 := StSelect + ' WHERE T_TIERS ="'+codeprospect+'"';
      QQ:=OpenSQL(StSelect2,TRUE) ;
      if Not QQ.EOF then TobProsp.detail[i].selectDB('',QQ,false);
      Ferme(QQ) ;
      end;
   L.ClearSelected;
   end;
FiniMove;

{$IFDEF EAGLCLIENT}
{AFAIREEAGL}
{$ELSE}
Ext := ExtractFileExt(Maquette);
if Ext <> '' then
   begin
   if Ext = '.doc' then
      ConvertDocFile(Maquette,DocGenere,Nil,Nil,Nil,OnGetVar,Nil,Nil)
   else
      ConvertRTFFile(Maquette,DocGenere,Nil,Nil,Nil,OnGetVar,Nil,Nil) ;
   end;
{$ENDIF}

TobProsp.free;
TobProsp:=Nil;
end ;


procedure TOF_EditionBI.OnGetVar( Sender: TObject; VarName: String; VarIndx: Integer;var Value: variant) ;
var TTOB : tob ;
BEGIN
if VarName='NBPROSPECT' then
   begin
      Value:=TobProsp.Detail.Count;
      exit;
   end;
if (Pos('T_',VarName)<= 0) and (Pos('RPR_',VarName)<=0) and (Pos('YTC_',VarName)<=0) and (Pos('C_',VarName)<=0) then exit;
TTob:=TobProsp.Detail[VarIndx-1];
if TTOB.GetValue(VarName) = NULL then
   begin
   if ((Pos('RPRLIBVAL',VarName)>0) or (Pos('VALLIBRE',VarName)>0)) then TTOB.PutValue(VarName,0)
   else if ((Pos('RPRLIBDATE',VarName)>0) or (Pos('DATELIBRE',VarName)>0)) then TTOB.PutValue(VarName,IDate1900)
   else TTOB.PutValue(VarName,'');
   end;
if Pos('BLOCNOTE',VarName)>0
      then  begin BlobToFile(VarName,TTob.GetValue(VarName)); Value:=''; end else
            if ChampToType( VarName)= 'COMBO' then Value:=RechDom(ChampToTT( VarName),TTob.GetValue(VarName),false)
               else Value:=TTOB.GetValue(VarName) ;
END ;

Function TOF_EditionBI.MailingValidDoc : Boolean;
var stDocWord , stMaquette: string;
begin

Result:=False;

stDocWord:=GetControlText('DOCGENERE');
stMaquette:=GetControlText('MAQUETTE');

//Verification si une maquette a été sélectionnée par l'utilisateur
if (stMaquette='') or (pos('\*.',stMaquette)<>0) then
   begin
   PGIBox('Vous devez choisir une maquette', ecran.caption);
   exit;
   end;

//Verification si la maquette sélectionnée existe
if not FileExists(stMaquette) then
   begin
   PGIBox('La maquette n''existe pas', ecran.caption);
   exit;
   end;

if (stDocWord='') then
   begin
   PGIBox('Vous devez mentionner un document', ecran.caption);
   exit;
   end;

if FileExists(stDocWord) then
   if PGIAsk('Ce fichier existe déjà, voulez-vous l''écraser ?',ecran.caption)=mryes then
   		DeleteFile(stDocWord )
   else
   		exit;

//gestion des documents par mail
if GetControlText('ENVOIMAIL')='X' then
   Begin
	 if GetControlText('OBJETMAIL') = '' then
   	  begin
   	  PGIBox('Vous devez renseigner l''objet du message', ecran.caption);
   		exit;
   		end;
	 if GetControlText('MAIL') <> 'X' then
	 		begin
   		PGIBox('Vous devez sélectionner les contacts avec e-mail (Onglet compléments)', ecran.caption);
   		exit;
   		end;
   end;

Result:=True;

end;


procedure TOF_EditionBI.ChangeMaquette;
var stMaquette: string;
    OkOk : Boolean ;
begin

OkOk:=False;

stMaquette:=getControlText('MAQUETTE');

if (stMaquette<>'') and (pos('\*.',stMaquette)=0) and  FileExists(stMaquette) then
   OkOk:=true;
SetControlEnabled('BMAQUETTE',  OkOk) ;

SetControlEnabled('BMAILINGDOC',  OkOk) ;

end;

procedure TOF_EditionBI.ChangeDocument;
var stDocument: string;
    OkOk : boolean ;
begin
OkOk:=False;
stDocument:=getControlText('DOCGENERE');
if (stDocument<>'') and (pos('\*.',stDocument)=0) and  FileExists(stDocument) then OkOk:=true;
SetControlEnabled('BDOCUMENT',OkOk) ;
end;

procedure TOF_EditionBI.Publipostage;
var StTitre,stDocument,stMaquette,ArgEmail: string;
    DS : THquery;
begin

stMaquette:=getControlText('MAQUETTE');
stDocument:=getControlText('DOCGENERE');

  if GetControlText('ENVOIMAIL') = 'X' then
  begin
    {$IFDEF EAGLCLIENT}
    if TFMul(Ecran).Fetchlestous then
    begin
    {$ENDIF}
    ArgEmail := 'EMAIL;-;ADRESSEEMAIL;';
    ArgEmail := ArgEmail + GetControlText('OBJETMAIL');
    DS:=THquery(TFMul(Ecran).Q) ;
    StTitre :=  TFMul(Ecran).Q.titres ;
    LancePublipostage(ArgEmail,stMaquette,stDocument,Nil,StTitre,DS,True);
    {$IFDEF EAGLCLIENT}
    end;
    {$ENDIF}
    exit;
  end;

{$IFDEF EAGLCLIENT}
if TFMul(Ecran).Fetchlestous then
   begin
{$ENDIF}
   DS:=THquery(TFMul(Ecran).Q) ;
   StTitre :=  TFMul(Ecran).Q.titres ;
   LancePublipostage('FILE',stMaquette,stDocument,Nil,StTitre,DS,True)
{$IFDEF EAGLCLIENT}
   end;
{$ENDIF}

end;


Procedure BTPEditionBI_ChangeMaquette( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFmul) then ToTof:=TFmul(F).LaTof else exit;
  TOF_EditionBI(totof).ChangeMaquette;
end;
Procedure BTPEditionBI_ChangeDocument( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFmul) then ToTof:=TFmul(F).LaTof else exit;
  TOF_EditionBI(totof).ChangeDocument;
end;

Procedure BTPPublipostage( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFmul) then ToTof:=TFmul(F).LaTof else exit;
  TOF_EditionBI(totof).Publipostage;
end;

Function BTPGetTitreGrid( parms: array of variant; nb: integer ) : variant ;
var  F : TForm ;
begin
  Result:='';
  F:=TForm(Longint(Parms[0])) ;
//  if (F is TFmul) then result:=getTitreGrid(TFmul(F).Fliste) else exit;
  if (F is TFmul) then result:=TFMul(F).Q.titres else exit;
end;


procedure TOF_EditionBI.MailingDoc(sender: TObject);
begin

if MailingValidDoc() then Publipostage();

ChangeDocument;

end;

procedure TOF_EditionBI.BOuvrir_Click(sender: TObject);
var StDocGenere	: String;
		stMaquette 	: String;
begin

if GetParamSoc('SO_EDITWORD') Then
  Begin
	stMaquette:=getControlText('MAQUETTE');
	StDocGenere:=getControlText('DOCGENERE');
	if MailingValidDoc() then EditionBI(stMaquette, StDocGenere);
	ChangeDocument ;
  end
else
	LanceEdition;

end;

//Procedure de lancement de l'édition en fonction du code Affaire sélectionné
procedure TOF_EditionBI.LanceEdition;
var F 			 	  : TFMul ;
    stWhere  		: sTring;
    stNature 		: string ;
	  Ind				  : Integer;
    //
	  NbExemplaire	: integer;
    NbPiece 		  : integer;
    stModele		  : String;
    NomPDF			  : String;
    TitreEtat 	  : String;
    TL 				    : TList;
    TT 				    : TStrings;
    BAp				    : Boolean;
    OldQRThread 	: boolean;
    ENomPDF     	: THEdit ;
    Pages 			  : TpageControl;
    TobAppel      : Tob;
    TOBL          : Tob;
    //
    S_PGI_QRPDF         : Boolean ;
    S_PGI_QRPDFQueue    : String;
    S_PGI_QRPDFMerge    : String;
    S_PGI_QRMultiThread :Boolean;
begin

  F:=TFMul(Ecran);

  if (F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
     begin
     if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
     PGIBox('Aucun élément sélectionné', F.caption) ;
     Exit;
     end;

   TobAppel := TOB.Create ('', Nil, -1) ;
   if TobAppel = nil then exit ;

   if F.FListe.AllSelected then
	    ChargeEditionAll(TobAppel)
   Else
      ChargeEdition(TobAppel);

   if TobAppel = nil then
      Begin
      TobAppel.free;
      exit;
      end;

   if TobAppel.detail.Count-1 > 2000 then
	    Begin
      PgiBox('Trop de documents sélectionnés. Veuiller réduire la sélection',F.Caption);
      TobAppel.free;
      exit;
      end;

   Pages:=Nil;

   //if assigned(GetControl('BAPERCU')) then BApercu := TCheckBox(GetControl ('BAPERCU')).Checked;
   //BDuplicata := TCheckBox(GetControl ('BDUPLICATA')).Checked;
   ENomPDF:=THEdit(GetControl('NOMPDF'));
   NomPDF:='' ;

   stWhere := '';
   stNature := TheNature;
   stmodele := TheModele;
//   stModele := GetParamsoc('SO_BTMODBI');
   NbPiece :=1;

   if ENomPDF <> Nil then NomPDF:=uppercase(ENomPDF.Text) ;

   S_PGI_QRPDF:=V_PGI.QRPDF ;
   S_PGI_QRPDFQueue:=V_PGI.QRPDFQueue ;
   S_PGI_QRPDFMerge:=V_PGI.QRPDFMerge ;
   S_PGI_QRMultiThread:=V_PGI.QRMultiThread ;

   if Not NomCorrect(NomPDF) then
      NomPDF:=''
   else
      BEGIN
      V_PGI.QRPDF:=True ;
      V_PGI.QRPDFQueue:=NomPDF ;
      V_PGI.QRPDFMerge:='' ;
      V_PGI.QRMultiThread:=False ;
      END;

   if NomPDF='' then
      BAp := ChkApercu.checked
   else
      BAp :=True ;

  for Ind :=0 to TobAppel.Detail.count-1 do
  begin
    TOBL := TobAppel.detail[Ind];
    stWhere := 'AFF_AFFAIRE="'  + Tobl.GetValue('AFF_AFFAIRE') + '"';
    TitreEtat :=  Tobl.GetValue('AFF_AFFAIRE1') + ' ' +
                  Tobl.GetValue('AFF_AFFAIRE2') + ' ' +
                  Tobl.GetValue('AFF_AFFAIRE3');
    if (Ind > 0)and (Bap = False) then V_PGI.NoPrintDialog := True;
    //
	  if not frealise then TheTitre := 'Intervention : ' + TitreEtat;
    //if OptionEdition.LanceImpression(stwhere, TOBL) < 0 then V_PGI.IoError:=oeUnknown;
//    LanceEtat('E','INT',OptionEdition.Modele,OptionEdition.Apercu,OptionEdition.ListeExport, OptionEdition.DeuxPages,Nil,trim (stWhere),TheTitre,OptionEdition.Duplicata);
    LanceEtat('E',TheNature,OptionEdition.Modele,BAp,OptionEdition.ListeExport, OptionEdition.DeuxPages,Nil,trim (stWhere),TheTitre,OptionEdition.Duplicata);
  end;

    BAp:=false;

		V_PGI.QRPDF:=S_PGI_QRPDF ;
    V_PGI.QRPDFQueue:=S_PGI_QRPDFQueue ;
    V_PGI.QRPDFMerge:=S_PGI_QRPDFMerge ;
    V_PGI.QRMultiThread:=S_PGI_QRMultiThread ;

    if F.FListe.AllSelected then
       F.FListe.AllSelected:=False
    else
       F.FListe.ClearSelected;

    F.bSelectAll.Down := False ;

end;

procedure TOF_EditionBI.ChargeEtat;
begin
  if not frealise then
  begin
  TheType   := 'E';
  TheNature := 'INT';
  //FS#751 - ESPACS : dans l'onglet édition, récupérer par défaut le modèle des paramètres société
  TheModele := GetParamSocSecur('SO_BTMODBI', '');

  TheTitre  := 'Edition Bon d''intervention';
  end else
  begin
    TheType   := 'E';
    TheNature := 'BRI';
    TheModele := 'BRI';
    TheTitre  := 'Edition Rapport d''intervention';
  end;


  OptionEdition := TOptionEdition.Create(TheType, TheNature, TheModele, TheTitre, '', ChkApercu.Checked, ChkReduire.Checked, True, False, Chkduplicata.Checked, TPageControl(Ecran.FindComponent('Pages')), fEtat);

  OptionEdition.first := True;
  OptionEdition.ChargeListeEtat(fEtat,Idef);

  if TheModele = '' then
  begin
    if fetat.itemindex  >= 0 then TheModele := FETAT.values[fetat.itemindex];
  end;
end;

Procedure TOF_EditionBI.OnChangeEtat(Sender : Tobject);
Begin

   OptionEdition.Modele := FETAT.values[fetat.itemindex];
   TheModele  := OptionEdition.Modele;

end;

//Chargement des enregistrements à éditer si selectedall
Procedure TOF_EditionBI.ChargeEditionAll(TobAppel : Tob);
Var	StWhere			: String;
    StArg				: String;
		QQ					: TQuery;
    F 					: TFMul ;
Begin

F:=TFMul(Ecran);

stWhere := RecupWhereCritere(F.Pages);

if stWhere = '' then
	stwhere := 'WHERE '
else
	stWhere := stWhere + ' AND ';

StArg := 'SELECT count(AFF_AFFAIRE) as nbr FROM (AFFAIRE ';
StArg := StArg + 'LEFT JOIN TIERS ON AFF_TIERS = T_TIERS) ';
StArg := StArg + 'LEFT JOIN RESSOURCE ON AFF_RESPONSABLE = ARS_RESSOURCE ';
StArg := StArg + 'LEFT JOIN ADRESSES ON AFF_AFFAIRE = ADR_REFCODE ';
StArg := StArg + stWhere + 'AFF_AFFAIRE0 = "W"';

QQ := OpenSQL (StArg, True);

if QQ.Eof then
	 Begin
	 ferme (QQ);
   exit;
   end;

If (QQ.FindField ('nbr').AsInteger > 2000) then
	 Begin
	 ferme (QQ);
   exit;
   end;

Ferme (QQ);

StArg := 'SELECT * FROM (AFFAIRE ';
StArg := StArg + 'LEFT JOIN TIERS ON AFF_TIERS = T_TIERS) ';
StArg := StArg + 'LEFT JOIN RESSOURCE ON AFF_RESPONSABLE = ARS_RESSOURCE ';
StArg := StArg + 'LEFT JOIN ADRESSES ON AFF_AFFAIRE = ADR_REFCODE ';
StArg := StArg + stWhere + 'AFF_AFFAIRE0 = "W"';

QQ := OpenSQL (StArg, True);

if not QQ.EOF then
   TobAppel.LoadDetailDB('APPEL', '', '', QQ, False)
else
   TobAppel := nil;

Ferme (QQ) ;

end;

//Chargement des enregistrements à éditer si selected a la main
Procedure TOF_EditionBI.ChargeEdition(TobAppel : Tob);
Var	StWhere			: String;
		QQ					: TQuery;
    F 					: TFMul ;
    Ind         : Integer;
begin

F:=TFMul(Ecran);

for Ind:=0 to F.FListe.NbSelected-1 do
    begin
    F.FListe.GotoLeBOOKMARK(Ind);
{$IFDEF EAGLCLIENT}
    F.Q.TQ.Seek(F.FListe.Row-1);
{$ENDIF}
    if STWhere <> '' then StWhere := StWhere + ' OR ';
    stWhere := StWhere + 'AFF_AFFAIRE="'  + F.Q.FindField('AFF_AFFAIRE').AsString + '"';
	  end;

QQ := OpenSql('SELECT * FROM AFFAIRE WHERE ' + stWhere , True);
if not QQ.EOF then
	 TobAppel.LoadDetailDB('APPEL', '', '', QQ, False)
else
	 TobAppel := nil;

Ferme (QQ);

end;

//Controle si le nom du Fichier PDF est valide.
Function TOF_EditionBI.NomCorrect (NomPDF : String) : Boolean ;
Var FFText 		: TextFile ;
BEGIN
Result:=False;

if NomPDF='' then Exit ;

if Pos('.PDF',NomPDF)<=0 then Exit ;

AssignFile(FFText,NomPDF) ;

{$i-} Rewrite(FFText) ; {$i+}

if IoResult=0 then
   BEGIN
   CloseFile(FFText) ;
   DeleteFile(NomPDF) ;
   Result:=True ;
   END ;

END ;

procedure TOF_EditionBI.DoubleClick(Sender: TObject);
var aff : string;
begin


	Aff := Grille.DataSource.DataSet.FindField('AFF_AFFAIRE').AsString ;

  AGLLanceFiche('BTP','BTAPPELINT','','','ACTION=MODIFICATION; CODEAPPEL:'+ aff + ';ETAT=AFF');

end;

Initialization
registerclasses([TOF_EditionBI]);
RegisterAglProc('EditionBI',TRUE,2,BTPEditionBI);
RegisterAglFunc('MailingValidDoc', TRUE , 0, BTPMailingValidDoc);
RegisterAglProc( 'ChangeMaquette', TRUE ,0 , BTPEditionBI_ChangeMaquette);
RegisterAglProc( 'ChangeDocument', TRUE , 0, BTPEditionBI_ChangeDocument);
RegisterAglFunc( 'GetTitreGrid', TRUE ,0 , BTPGetTitreGrid);
RegisterAglProc( 'Publipostage', TRUE ,0 , BTPPublipostage);
end.

