unit AssistExport;

interface

uses Windows,
     Messages,
     SysUtils,
     Classes,
     Graphics,
     Controls,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}
     dbtables,
     {$ELSE}
     uDbxDataSet,
     {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
     Forms,
     Dialogs,
  	 Assist,
     HSysMenu,
     hmsgbox,
     StdCtrls,
     ComCtrls,
     ExtCtrls,
     HPanel,
     HTB97,
  	 Hctrls,
     HFLabel,
     UTOB,
     HEnt1,
     ParamSoc,
     UtilInterComsX,
     Mask;

type
  TFAssistExport = class(TFAssist)
    PTitre						: TPanel;
    PRESUME						: TPanel;
		//
    TGENERALITE				: TTabSheet;
    TENVOIFIL					: TTabSheet;
    TRESUME						: TTabSheet;
    //
    ChkExport					: TCheckBox;
    //
    LblDu							: TLabel;
    LblAu							: TLabel;
    LblTraitement			: TLabel;
    LblPeriode				: TLabel;
    LblFile						:	TLabel;
    LblMessagerie			: TLabel;
    //
    DateDeb						: TMaskEdit;
    DateFin						: TMaskEdit;
    //
    ChkEnvoiEcriture	: TRadioButton;
    //
    GroupBox2					: TGroupBox;
    GroupBox5					: TGroupBox;
    //
    HLabel2						: THLabel;
    HLabel1						: THLabel;
    HLabel7						: THLabel;
    HLabel8						: THLabel;
    Panel1						: TPanel;
    ChkMail						: TRadioButton;
    HLabel4						: THLabel;
    GroupBox3					: TGroupBox;
    HLabel3						: THLabel;
    ChkFile						: TRadioButton;
    PENVOIFIL					: TPanel;
    LblExchangeFile		: TLabel;
    FileTrF						: THCritMaskEdit;
    PENVOIMAIL				: TPanel;
    LblAdrMail				: TLabel;
    AdrMail						: TEdit;
    ChkMailImportance	: TCheckBox;
    LblCorpMail				: TLabel;
    CorpMessage				: TMemo;
    LblAttention			: TLabel;
    HLabel6						: THLabel;
    ODGetInfosTRA			: TOpenDialog;
    BlocNote					: TMemo;
    //
    procedure ChargeExport;
    procedure FormShow(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure ChkMailClick(Sender: TObject);
    procedure ChkFileClick(Sender: TObject);
    procedure FileTrFElipsisClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    //

  private
		{ Déclarations privées }
    Frequence  : String;
    OK_Fichier : Boolean;
    Ok_Mail    : Boolean;
    ChargeComsX : ChgComsX;
    procedure AffichePControl;
    procedure ChargelastPage;
    function AxevalidePourCptaDest(TypeLien: string): Boolean;
    //
  public
    { Déclarations publiques }
  end;

  Procedure AppelAssistExport;

  Const
  //
  // libellés des messages
  TexteMessage: array[ 1..4 ] of hstring = (
    {1}  'La date de début d''export est incorrecte.'
    {2}, 'La date de fin d''export est incorrecte.'
    {3}, 'Veuillez renseigner un fichier d''export SVP.'
    {4}, 'Veuillez renseigner une adresse mail SVP.'
     );


implementation


{$R *.DFM}

Procedure AppelAssistExport;
	var XX : TFAssistExport;
begin
	XX := TFAssistExport.Create (application);
  XX.ShowModal;
  XX.free;
end;

procedure TFAssistExport.ChargeExport;
Var JJ			: Word;
    MM 			: Word;
    AA			: Word;
begin

  BFin.Visible := False;
  //
  //ChkEnvoiEcriture.Enabled := False;

  //
  FileTrf.Text := GetParamSocSecur('SO_CPRDREPERTOIRE', 'C:\');
  if FileTrf.Text = '' then
     FileTrf.Text := 'C:\' + 'PG' + V_PGI.NoDossier + '.TRA'
  else
     FileTrf.Text := FileTrf.Text + '\' + 'PG' + V_PGI.NoDossier + '.TRA';
  //
  AdrMail.text := GetparamSocSecur('SO_CPRDEMAILCLIENT','');
  //
  Frequence := GetParamSocSecur('SO_CPFREQUENCESX','HEB');
  //
	AffichePControl;
  //
  //Chargement des dates d'écritures...
  if Frequence = 'HEB' Then
	   begin
     DecodeDate (V_PGI.dateEntree,AA,MM,JJ);
     DateDeb.text := DateToStr(PremierJourSemaine (NumSemaine(V_PGI.DateEntree),AA));
     DateFin.Text := DateToStr(PlusDate(StrToDate(DateDeb.Text),7,'J'));
     end
  else if Pos(Frequence ,'JOU;PON;')>0then
     begin
     DateDeb.Text := DateToStr(V_PGI.dateEntree);
     DateFin.Text := DateToStr(V_PGI.dateEntree);
     end
  else if Frequence ='MEN' then
	   begin
     DateDeb.Text := DateToStr(DebutDeMois(V_PGI.DateEntree));
     Datefin.Text := DateToStr(FinDeMois(V_PGI.DateEntree));
     end;

end;

procedure TFAssistExport.FormShow(Sender: TObject);
begin
  inherited;
  ChargeExport;
end;

procedure TFAssistExport.bSuivantClick(Sender: TObject);
begin
  inherited;

  bPrecedent.Enabled := True;

  //gestion dispo ou non dispo des bouton
	if P.ActivePageIndex = P.PageCount-1 then
  	 begin
     ChargelastPage;
     end;

end;

procedure TFAssistExport.bPrecedentClick(Sender: TObject);
begin
  inherited;

  bfin.Visible := false;
 	bSuivant.Enabled := True;

	if P.ActivePageIndex = 0 then
  	 bPrecedent.Enabled := false;

end;

procedure TFAssistExport.ChkFileClick(Sender: TObject);
begin
  inherited;

	AffichePControl;

end;

procedure TFAssistExport.ChkMailClick(Sender: TObject);
begin
  inherited;

	AffichePControl;

end;

Procedure TFAssistExport.AffichePControl;
Begin

	Ok_Fichier := ChkFile.checked;
  Ok_Mail := not ChkFile.checked;

  PENVOIFIL.Visible := Ok_Fichier;
  PENVOIMAIL.Visible := Ok_Mail;
  Hlabel4.Visible := Ok_Fichier;
  Hlabel6.Visible := Ok_mAIL;

end;

Procedure TFAssistExport.ChargelastPage;
Begin

  bfin.Visible := true;
  bSuivant.Enabled := false;

  LblTraitement.Caption := 'Traitement : Export des ecritures ';
  if ChkExport.Checked then
     LblTraitement.Caption := LblTraitement.Caption + '(intégration de celles déjà traitées)';

  LblPeriode.Caption		:= 'Période : ' + DateDeb.Text + ' au ' + DateFin.Text;

  if Ok_Fichier then
	   LblFile.Caption				:= 'Fichier : ' +  FileTrF.Text
  else
     LblMessagerie.caption	:= 'Messagerie : Envoie à l''adresse ' +  AdrMail.text;

  if ChkMailImportance.Checked then
       LblMessagerie.caption :=   LblMessagerie.caption + ' en Importance Haute';

end;

procedure TFAssistExport.FileTrFElipsisClick(Sender: TObject);
Var Extension : String;
    X : Integer;
begin
  inherited;

  if fileTrf.Text = '' then
	   ODGetInfosTRA.InitialDir := GetParamSocSecur('SOC_DATA', 'C:\')
  else
  	 ODGetInfosTRA.InitialDir := fileTrf.Text;

  if ODGetInfosTRA.Execute then
  	 begin
     if ODGetInfosTRA.FileName <> '' then
    		begin
        FileTrf.Text := ODGetInfosTRA.FileName;
        X := pos ('.', FileTrf.Text) ;
        if X = 0 then
           FileTrf.Text := ODGetInfosTRA.FileName + '.TRA'
        else
           Begin
	         Extension := Copy(FileTrf.Text, X + 1, 4);
           if Extension <> 'TRA' then
           	  FileTrf.Text := ODGetInfosTRA.FileName + '.TRA'
           else
              FileTrf.Text := ODGetInfosTRA.FileName;
   		     end;
        end;
  	 end;

end;

procedure TFAssistExport.bFinClick(Sender: TObject);
Var BlocNotes 	: TStringList;
		TransfertVert : string;
begin
  inherited;

  if BFin.Caption = 'Fin' Then
     Begin
     Close;
     Exit;
     end;

  BlocNote.Visible := False;
  BlocNote.Text := '';

  if not AxevalidePourCptaDest (GetParamSocSecur( 'SO_BTLIENCPTAS1', 'AUC' )) then
  begin
  	Exit;
  end;

  if not IsValidDate(DateDeb.text) then
     Begin
  	 exit;
     end;

  if not IsValidDate(DateFin.text) then
     Begin
  	 exit;
     end;

  if Ok_Fichier then
     Begin
     if FileTrF.Text = '' then
        Begin
        Exit;
        end;
     end;

  if Ok_Mail then
     Begin
     if AdrMail.text = '' then
        Begin
        exit;
        end;
     end;

  If PgiAsk(TraduireMemoire ('Validez-vous le traitement d''envoi des écritures ?')) = MrNo then Exit;

  if ChkExport.State = cbchecked then
     ChargeComsX.Envoye:='TRUE'
  else if ChkExport.State = cbUnchecked then
     ChargeComsX.Envoye:='FALSE'
  else
     ChargeComsX.Envoye := '';

  BlocNotes := TStringList.Create;

  TransfertVert := GetParamSocSecur( 'SO_BTLIENCPTAS1', 'AUC' );
  if pos (TransfertVert, 'S3;S5;QUA;WIN;')> 0 then TransfertVert := 'S5';

  //Chargement des zones nécessaire pour le transfert via ComsX
  With ChargeComsX Do
       Begin
       FichierEcr 		:= FileTrF.Text;
    	 GestEtab				:= 'FALSE';
       TransfertVers 	:= TransfertVert;
       Format 				:= 'ETE';
       DateArret			:= '';
       Exercice				:= '';
  		 DateDebut  		:= DateDeb.Text;
  		 DateDeFin  		:= DateFin.Text;
     	 TypeEcrRecup 	:= 'N';
 		   TypeEcrGen 		:= 'N';
       EMail 					:= AdrMail.Text; //Envoie rapport après traitement
	     Journal 				:= '[' + GetParamSocSecur( 'SO_EXPJRNX', '' ) + ']';
   		 FichierRapport := ExtractFileDir(FichierEcr) + '\RAPPORT.TXT';
       Societe 				:= V_PGI.CurrentAlias;
       User						:= V_PGI.UserLogin;
		   DateFile   		:= DateToStr(Now);
 		   TypeEcr 	  		:= '';
       TypeEcriture   := 'EXPORT';
       NatureTransfert:= 'JRL';
       end;

  //UnEchangeComSx.ExecuteAction (TmaSxExport,FileTrF.Text,Datedeb.Text,DateDeb.text,DateFin.text,EtatEnvoye);
  GenereFile(ChargeComsX,BlocNotes);

  If PgiAsk(TraduireMemoire ('Voulez-vous afficher le rapport ?')) = MrYes then
     Begin
     BlocNote.Visible := True;
     BlocNote.Text := BlocNotes.text;
     end;

  BlocNotes.Free;

  BFin.Caption := 'Fin';

end;

function TFAssistExport.AxevalidePourCptaDest (TypeLien : string): Boolean;

	function ComptageAxe : Integer;
  var QQ : TQuery;
  		TOBT : TOB;
      II : Integer;
  begin
    Result := 0;
		TOBT := TOB.Create ('LESA XES',nil,-1);
    QQ := OpenSQL('SELECT * FROM AXE',True,-1,'',true);
    if not QQ.Eof then TOBT.LoadDetailDB('AXE','','',QQ,false);
    ferme (QQ);
    for II := 0 to TOBT.detail.count -1 do
    begin
			if (TOBT.detail[iI].GetString('X_LIBELLE')<>'Axe N°'+InttoStr(II)) or
         (TOBT.detail[iI].GetString('X_SECTIONATTENTE')<>'') then Inc(Result) else Break;
    end;
    TOBT.Free;
  end;
  
begin
  Result := True;
	if TypeLien <> 'S1' then Exit;
  if TypeLien = 'S1' then if ComptageAxe > 1 then
  begin
    PGIError('Vous ne pouvez pas exporter des écritures ayant plus d''un axe pour LINE');
  	Result := false;
  end;
end;

end.
