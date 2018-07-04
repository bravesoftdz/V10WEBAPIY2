unit BTAssistImport;

interface

uses  Windows,
			Messages,
      SysUtils,
      Classes,
      Graphics,
      Controls,
      Forms,
      Dialogs,
  		Assist,
      StdCtrls,
      Mask,
      Hctrls,
      ComCtrls,
      HSysMenu,
      hmsgbox,
      ExtCtrls,
  	  HPanel,
      ParamSoc,
      HTB97,
      HEnt1,
		  UtilInterComsX;

type
  TFBTAssistImport = class(TFAssist)
    TRECEPTION			: TTabSheet;
    TRESUME					: TTabSheet;
    HLabel3					: THLabel;
    GroupBox3				: TGroupBox;
    HLabel6					: THLabel;
    ChkFile					: TRadioButton;
    LblExchangeFile	: TLabel;
    FileTrF					: THCritMaskEdit;
    ODGetInfosTRA		: TOpenDialog;
    PRESUME					: TPanel;
    LblTraitement		: TLabel;
    LblPeriode			: TLabel;
    LblFile					: TLabel;
    HLabel8					: THLabel;
    GroupBox5				: TGroupBox;
    HLabel7					: THLabel;
    BlocNote				: TMemo;
    PTitre					: TPanel;
    //
    procedure ChargeImport;
    procedure bPrecedentClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FileTrFElipsisClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);

  private
    { Déclarations privées }
    ChargeComsX : ChgComsX;
    //
    procedure ChargelastPage;
  public
    { Déclarations publiques }
  end;
  //
  Procedure AppelAssistImport;
  //
  Const
  // libellés des messages
  TexteMessage: array[ 1..1 ] of hstring = (
    {1}  'Le fichier de transfert n''existe pas !'
     );


implementation

{$R *.DFM}

Procedure AppelAssistImport;
	var XX : TFBTAssistImport;
begin
	XX := TFBTAssistImport.Create (application);
  XX.ShowModal;
  XX.free;
end;

procedure TFBTAssistImport.ChargeImport;
begin

	P.ActivePageIndex := 0;

  BFin.Visible := False;

  FileTrf.Text := GetParamSocSecur('SO_CPRDREPERTOIRE', 'C:\');
  if FileTrf.Text = '' then
     FileTrf.Text := 'C:\' //+ 'PG' + V_PGI.NoDossier + '.TRA'
  else
     FileTrf.Text := FileTrf.Text + '\' //+ 'PG' + V_PGI.NoDossier + '.TRA';

end;

procedure TFBTAssistImport.bSuivantClick(Sender: TObject);
begin
  inherited;

  bPrecedent.Enabled := True;

  if P.ActivePageIndex  = 1 then
     begin
     if Not FileExists(FileTrf.Text) then
        begin
        PGIBox(TexteMessage[1],'"Réception');
        P.ActivePageIndex := 0;
        exit;
        end;
     end;

  //gestion dispo ou non dispo des bouton
	if P.ActivePageIndex = P.PageCount-1 then
  	 begin
     ChargelastPage;
     end;

end;

procedure TFBTAssistImport.bPrecedentClick(Sender: TObject);
begin
  inherited;

  bfin.Visible := false;
 	bSuivant.Enabled := True;

	if P.ActivePageIndex = 0 then
  	 bPrecedent.Enabled := false;

end;

Procedure TFBTAssistImport.ChargelastPage;
Begin

  bfin.Visible := true;
  bSuivant.Enabled := false;

  LblTraitement.Caption := 'Traitement : Réception des ecritures ';

  LblPeriode.Caption		:= 'Période : Général';

  LblFile.Caption				:= 'Fichier : ' +  FileTrF.Text
 
end;

procedure TFBTAssistImport.FormShow(Sender: TObject);
begin
  inherited;
  ChargeImport;
end;

procedure TFBTAssistImport.FileTrFElipsisClick(Sender: TObject);
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

procedure TFBTAssistImport.bFinClick(Sender: TObject);
Var BlocNotes 	: TStringList;
begin
  inherited;

  if BFin.caption = 'Fin' then
     Begin
     Close;
     Exit;
     end;

  BlocNote.Visible := False;
  BlocNote.Text := '';

  If PgiAsk(TraduireMemoire ('Validez-vous la récupération des écritures ?')) = MrNo then Exit;

  BlocNotes := TStringList.Create;

  //Chargement des zones nécessaire pour le transfert via ComsX
  With ChargeComsX Do
       Begin
       FichierEcr 		:= FileTrF.Text;
    	 GestEtab				:= 'FALSE';
       TransfertVers 	:= 'S1';
       Format 				:= 'ETE';
       DateArret			:= '';
       Exercice				:= '';
  		 DateDebut  		:= DateTimeToStr(idate1900);
  		 DateDeFin  		:= DateTimeToStr(IDate2099);
     	 TypeEcrRecup 	:= 'N';
 		   TypeEcrGen 		:= 'N';
       EMail 					:= ''; //Envoie rapport après traitement
	     Journal 				:= '[' + GetParamSocSecur( 'SO_EXPJRNX', '' ) + ']';
   		 FichierRapport := ExtractFileDir(FichierEcr) + 'RAPPORT.TXT';
       Societe 				:= V_PGI.CurrentAlias;
       User						:= V_PGI.UserLogin;
		   DateFile   		:= DateToStr(Now);
 		   TypeEcr 	  		:= '';
       TypeEcriture   := 'IMPORT';
       NatureTransfert:= 'JRL';
       end;

  //UnEchangeComSx.ExecuteAction (TmaSxExport,FileTrF.Text,Datedeb.Text,DateDeb.text,DateFin.text,EtatEnvoye);
  ImportFile(ChargeComsX,BlocNotes);

  If PgiAsk(TraduireMemoire ('Voulez-vous afficher le rapport ?')) = MrYes then
     Begin
     BlocNote.Visible := True;
     BlocNote.Text := BlocNotes.text;
     end;

  BlocNotes.Free;

  //Mise à jour automatique des tiers client en Facture HT au lieu de TTC
  ExecuteSQL('UPDATE TIERS SET T_FACTUREHT="X"');

  BFin.Caption := 'Fin';

end;

end.
