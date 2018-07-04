unit GcTraitCli;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
{$IFDEF EAGLCLIENT}
{$ELSE}
  dbTables,
{$ENDIF}
  Hgauge,TiersUtil, UTob,HPanel,UiUtil, Buttons,ParamSoc,M3FP;

  Procedure TraitClientSerie ;

type
  TFTraitcli = class(TFAssist)
    TabSheet1: TTabSheet;
    infomess: TPanel;
    infomess2: TLabel;
    infomess1: TLabel;
    nbclitrait: TEnhancedGauge;
    TabSheet2: TTabSheet;
    GRPART: TGroupBox;
    LIB1: TLabel;
    LIB2: TLabel;
    LIB3: TLabel;
    FMTTIERSTAILLEMOT: TEdit;
    GRENT: TGroupBox;
    LIB4: TLabel;
    LIB5: TLabel;
    GRADR: TGroupBox;
    LIB6: TLabel;
    LIB7: TLabel;
    LIB8: TLabel;
    FMTTIERSSUPESPACE: TCheckBox;
    BOuvrir: TToolbarButton97;
    ToolbarButton971: TToolbarButton97;
    FMTTIERSLIBELLE: THValComboBox;
    FMTTIERSPARTIC: THValComboBox;
    FMTTIERSPRENOM: THValComboBox;
    FMTSOCLIBELLE: THValComboBox;
    FMTSOCPRENOM: THValComboBox;
    FMTTIERSADR1: THValComboBox;
    FMTTIERSADR2: THValComboBox;
    FMTTIERSADR3: THValComboBox;
    FMTTIERSVILLE: THValComboBox;
    LIB9: TLabel;
    bmajparamz: TButton;
    Panel1: TPanel;
    RBtous: TRadioButton;
    RBclientpart: TRadioButton;
    RBentreprise: TRadioButton;
    lancetrait: TBitBtn;
    BStop: TBitBtn;
    procedure BStopClick(Sender: TObject);
    procedure lancetraitClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    function GetPageName : String ;
    function GetPage : TTabSheet ;
    procedure MajParamClick(Sender: TObject);
    procedure activebmaj(Sender: TObject);
  private
    nbcli, nbtrait, nbclimaj: integer;
    ArretUser, fairemaj : Boolean ;
    { Déclarations privées }
	 procedure FormatetousClients ;
	 Function  FormateUnClient( TOBCli : TOB ) : Integer ;
  public
    { Déclarations publiques }
  end;

const
	// Nom des champs à formater
	ChpAFormater: array[1..7] of string 	= (
          {1}        'T_LIBELLE'
          {2}       ,'T_PRENOM'
          {3}       ,'T_ABREGE'
          {4}       ,'T_ADRESSE1'
          {5}       ,'T_ADRESSE2'
          {6}       ,'T_ADRESSE3'
          {7}       ,'T_VILLE'
                     );

implementation

{$R *.DFM}
Procedure TraitClientSerie ;
var  Ftraitcli: TFtraitcli;
BEGIN
	Ftraitcli:=TFtraitcli.Create(Application) ;
   Try
       Ftraitcli.ShowModal;
   Finally
       Ftraitcli.free;
   End;
END ;

procedure TFtraitcli.lancetraitClick(Sender: TObject);
begin
  inherited;
	ArretUser := False ;
   nbclitrait.Progress := 0;
   if HShowMessage('0;Confirmation;Confirmez vous le lancement du traitement ?;Q;YN;N;N;','','')<>mrYes then exit;
	BStop.Enabled := True;
	lancetrait.Enabled := False;
	nbclitrait.Visible := True;
   nbclimaj := 0;
   fairemaj := False;
	FormatetousClients;
	nbclitrait.Visible := False;
end;


///////////////////////////////////////////////////////////////////////////////////////
//  MajDesClient : mise à jour de l'ensemble des clients
///////////////////////////////////////////////////////////////////////////////////////
procedure TFtraitcli.FormatetousClients ;
Var TOBClient, TOBD : TOB ;
    QQ        : TQuery ;
    i_ind : integer ;
BEGIN
// Création de la TOB mère
TOBClient := TOB.Create('Clients', Nil, -1) ;
// Chargement des clients
QQ:=Nil;
if RBclientpart.Checked then
	QQ:=OpenSQL('SELECT T_AUXILIAIRE, T_PARTICULIER, T_LIBELLE, T_PRENOM,T_ABREGE,T_ADRESSE1,T_ADRESSE2'+
               ',T_ADRESSE3,T_VILLE FROM TIERS WHERE T_NATUREAUXI = "CLI"'+
               ' AND T_PARTICULIER = "X"', True)
else if RBentreprise.Checked then
	QQ:=OpenSQL('SELECT T_AUXILIAIRE, T_PARTICULIER, T_LIBELLE, T_PRENOM,T_ABREGE,T_ADRESSE1,T_ADRESSE2'+
               ',T_ADRESSE3,T_VILLE FROM TIERS WHERE T_NATUREAUXI = "CLI"'+
               ' AND T_PARTICULIER = "-"', True)
else if RBtous.Checked then
	QQ:=OpenSQL('SELECT T_AUXILIAIRE, T_PARTICULIER, T_LIBELLE, T_PRENOM,T_ABREGE,T_ADRESSE1,T_ADRESSE2'+
               ',T_ADRESSE3,T_VILLE FROM TIERS WHERE T_NATUREAUXI = "CLI"', True) ;
TOBClient.LoadDetailDB('TIERS', '', '', QQ, False, True);
if QQ<>Nil then Ferme(QQ) ;
nbcli := TOBClient.Detail.Count;
nbclitrait.MaxValue := nbcli;
// Parcours des clients non fermé
for i_ind := 0 to nbcli -1 do
    BEGIN
    TOBD := TOBClient.Detail[i_ind] ;
    FormateUnClient (TOBD) ;
    END;
if ArretUser then
begin
	MessageAlerte('Arrêt demandé par l''utilisateur,'+#10#13+' aucune mise à jour effectuée !');
end else
begin
   if (nbclimaj > 0) then
   begin
    TOBClient.UpdateDB(True) ;
   end;
	MessageAlerte('Traitement terminé, Nombre de clients traités : '+InttoStr(nbclimaj));
end;
// Libére la liste des clients
TOBClient.Free ;
BStop.Enabled := False;
lancetrait.Enabled := True;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  MajUnClient : traitement d'une fiche client
///////////////////////////////////////////////////////////////////////////////////////
Function TFtraitcli.FormateUnClient( TOBCli : TOB ) : Integer ;
var texte : string;
    Ind : integer ;
    clipart : boolean;
begin
  fairemaj := False;
  if ArretUser then
  begin
      Result := (-1);
      Exit ;
  end else
  begin
	  nbtrait := nbtrait + 1;
	  nbclitrait.Progress := nbclitrait.progress + 1;
	  Result := 0 ;
	  if TOBCli.GetValue('T_PARTICULIER') = 'X' then
	    	clipart := True
	  else
	   	clipart := False;

	  for Ind := Low(ChpAFormater) to High(ChpAFormater) do
	  Begin
	   	texte := VerifAppliqueFormat(ChpAFormater[Ind], TOBCli.GetValue (ChpAFormater[Ind]), clipart);
	   	if (texte <> TOBCli.GetValue (ChpAFormater[Ind])) then
         begin
	   		TOBCli.PutValue(ChpAFormater[Ind], texte);
			   fairemaj := True;
         end;
	  End ;
     if (fairemaj = True) then
     begin
		  nbclimaj := nbclimaj + 1;
     end;
  End;
end;



procedure TFtraitcli.FormShow(Sender: TObject);
begin
  inherited;
ArretUser := False ;
bFin.Visible := True;
bFin.Enabled := True;
BStop.Enabled := False;
lancetrait.Enabled := True;
nbclitrait.Visible := False;
FMTTIERSLIBELLE.value := varAsType(GetParamSoc('SO_FMTTIERSLIBELLE'), varString);
FMTTIERSPRENOM.value := varAsType(GetParamSoc('SO_FMTTIERSPRENOM'), varString);
FMTTIERSADR1.value := varAsType(GetParamSoc('SO_FMTTIERSADR1'), varString);
FMTTIERSADR2.value := varAsType(GetParamSoc('SO_FMTTIERSADR2'), varString);
FMTTIERSADR3.value := varAsType(GetParamSoc('SO_FMTTIERSADR3'), varString);
FMTTIERSVILLE.value := varAsType(GetParamSoc('SO_FMTTIERSVILLE'), varString);
FMTSOCLIBELLE.value := varAsType(GetParamSoc('SO_FMTSOCLIBELLE'), varString);
FMTSOCPRENOM.value := varAsType(GetParamSoc('SO_FMTSOCPRENOM'), varString);
FMTTIERSPARTIC.value := varAsType(GetParamSoc('SO_FMTTIERSPARTIC'), varString);
FMTTIERSTAILLEMOT.text := varAsType(GetParamSoc('SO_FMTTIERSTAILLEMOT'), varInteger);
FMTTIERSSUPESPACE.checked := varAsType(GetParamSoc('SO_FMTTIERSSUPESPACE'), varBoolean);
bmajparamz.Enabled := False;
end;

procedure TFtraitcli.bFinClick(Sender: TObject);
begin
  inherited;
Close ;
end;

procedure TFtraitcli.BStopClick(Sender: TObject);
begin
   if HShowMessage('0;Confirmation;Confirmez vous l''arrêt du traitement ?;Q;YN;N;N;','','')<>mrYes then
   begin
     ArretUser := False;
	  exit;
	end;
   ArretUser := True ;
end;

function TFtraitcli.GetPageName : String ;
BEGIN
result := P.ActivePage.Caption ;
END ;

function TFtraitcli.GetPage : TTabSheet ;
BEGIN
result := P.ActivePage ;
END ;

procedure TFTraitcli.MajParamClick(Sender: TObject);
begin
  inherited;
  // modification des paramètres société
SetParamsoc('SO_FMTTIERSLIBELLE',FMTTIERSLIBELLE.value) ;
SetParamsoc('SO_FMTTIERSPRENOM',FMTTIERSPRENOM.value) ;
SetParamsoc('SO_FMTTIERSADR1',FMTTIERSADR1.value) ;
SetParamsoc('SO_FMTTIERSADR2',FMTTIERSADR2.value) ;
SetParamsoc('SO_FMTTIERSADR3',FMTTIERSADR3.value) ;
SetParamsoc('SO_FMTTIERSVILLE',FMTTIERSVILLE.value) ;
SetParamsoc('SO_FMTSOCLIBELLE',FMTSOCLIBELLE.value) ;
SetParamsoc('SO_FMTSOCPRENOM',FMTSOCPRENOM.value) ;
SetParamsoc('SO_FMTTIERSPARTIC',FMTTIERSPARTIC.value) ;
SetParamsoc('SO_FMTTIERSTAILLEMOT',FMTTIERSTAILLEMOT.text) ;
SetParamsoc('SO_FMTTIERSSUPESPACE',FMTTIERSSUPESPACE.checked) ;
bmajparamz.Enabled := False;
end;

procedure TFTraitcli.activebmaj(Sender: TObject);
begin
  inherited;
 // activation du bouton de mémorisation
bmajparamz.Enabled := True;
end;

Procedure AGLTraitClientSerie ( Parms : array of variant ; nb : integer ) ;
BEGIN
TraitClientSerie ;
END ;

Initialization
RegisterAglProc('TraitClientSerie',False,0,AGLTraitClientSerie) ;

end.


