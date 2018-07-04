unit RevDist ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, ExtCtrls, StdCtrls, Mask, Hctrls, ComCtrls, HSysMenu, hmsgbox,
  HTB97, HPanel, ParamDat, Hent1, UIUtil, S1Util, DB, Hqry, HStatus, uTob, ParamSoc,
  dbtables, MailOL ;

procedure RevisionAdistance ( Expert : boolean ) ;

type
  TFRevDist = class(TFAssist)
    Choix: TTabSheet;
    Periode: TTabSheet;
    Fichier: TTabSheet;
    Mail: TTabSheet;
    rg: TRadioGroup;
    FArrete: TCheckBox;
    Label3: TLabel;
    edPath: THCritMaskEdit;
    FMail: TCheckBox;
    Label2: TLabel;
    FEMail: THCritMaskEdit;
    FMode: TLabel;
    Bevel1: TBevel;
    Label5: TLabel;
    Label7: TLabel;
    Bevel2: TBevel;
    Label4: TLabel;
    Bevel3: TBevel;
    Bevel4: TBevel;
    Label8: TLabel;
    Label10: TLabel;
    Bevel5: TBevel;
    Label11: TLabel;
    FCorpsMail: TMemo;
    Label12: TLabel;
    Resume: TTabSheet;
    Bevel6: TBevel;
    Label13: TLabel;
    Label1: TLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    Label9: TLabel;
    Label15: TLabel;
    FLib1: TLabel;
    FVal1: TLabel;
    CM: TOpenDialog;
    FLib2: TLabel;
    FVal2: TLabel;
    FLib3: TLabel;
    FVal3: TLabel;
    FLib4: TLabel;
    FVal4: TLabel;
    Bevel7: TBevel;
    FAttClient: TLabel;
    bExec: TToolbarButton97;
    FHigh: TCheckBox;
    HDiv: THMsgBox;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure rgClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bExecClick(Sender: TObject);
    procedure edPathElipsisClick(Sender: TObject);
    procedure PChange(Sender: TObject); override ;
    procedure FMailClick(Sender: TObject);
  private
    Expert : boolean;
    procedure Recevoir;
    procedure Envoyer;
    procedure ExportPiece ( Fichier : String ; NumeroPiece : Integer; Ajout,SansPerimetre : Boolean ; NumPaquet : integer ) ;
    procedure ExportPerimetreLettrage( TOBEcr : TOB ; Fichier : string ; Ajout : Boolean ; NumPaquet : integer ) ;
    procedure AnnulerEnvoi ;
  public
    function  PreviousPage : TTabSheet ; Override ;
    function  NextPage : TTabSheet ; Override ;
  end;

implementation

uses cloM;

{$R *.DFM}

procedure RevisionAdistance ( Expert : boolean ) ;
Var X : TFRevDist;
    PP : THPanel ;
    j,m,a: word;
    LaDate : TDateTime;
begin
SourisSablier;
X:=TFRevDist.Create(Application) ;
X.Expert:=Expert ;
PP:=FindInsidePanel ;
If PP=Nil then
   BEGIN
    Try
     X.ShowModal;
    finally
     X.Free ;
    end ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
SourisNormale ;
end;

procedure TFRevDist.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
if IsInside(Self) then Action:=caFree ;
end;

procedure TFRevDist.rgClick(Sender: TObject);
begin
case Rg.ItemIndex of
   0 : bExec.Caption:=HDiv.Mess[0] ;
   1 : bExec.Caption:=HDiv.Mess[1] ;
   2 : bExec.Caption:=HDiv.Mess[2] ;
   end;
FAttClient.Visible:=(Rg.ItemIndex=0) and (not Expert) ;
end;

procedure TFRevDist.PositionneDate ;
Var a,m,j : Word ;
    LaDate : TDateTime ;
BEGIN
DecodeDate(Date,a,m,j) ;
if m>1 then m:=m-1 else BEGIN m:=12 ; a:=a-1 ; END ;
LaDate:=EncodeDate(a,m,j) ;
E_DATECOMPTABLE.Text:=DateToStr(FindeMois(LaDate)) ;
END ;

procedure TFRevDist.RempliCorps ;
BEGIN
FCorpsMail.Lines.Clear ;
s:=HDiv.Mess[7] ;
s:=s+' '+V_PGI.NomSociete+' '+MSG.Mess[8] ;
if V_PGI.LaSerie=S5 then s:=s+' S5' else s:=s+' S7' ;
FCorpsMail.Lines.Add(s) ;
FCorpsMail.Lines.Add('') ;
FCorpsMail.Lines.Add(HDiv.Mess[9]) ;
FCorpsMail.Lines.Add('') ;
FCorpsMail.Lines.Add(HDiv.Mess[10]) ;
FCorpsMail.Lines.Add('') ;
if Expert then FCorpsMail.Lines.Add(GetParamSocSecur('SO_CABINET',''))
          else FCorpsMail.Lines.Add(GetParamSocSecur('SO_LIBELLE','')) ;
END ;

procedure TFRevDist.RenseigneRG ;
BEGIN
if Expert then
   BEGIN
   Rg.Items[0]:=HDiv.Mess[5] ;
   Rg.Items[1]:=HDiv.Mess[6] ;
   END ;
RgClick(nil) ;
END ;

procedure TFRevDist.ChargeparamsCabinet ;
BEGIN
if Expert then FMode.Caption:=HDiv.Mess[3] ;
if Not Expert then Rg.Items.Add(HDiv.Mess[4]) ;
if Expert then FEMail.Text:=GetParamSocSecur('SO_MAIL','')) else FEMail.Text:=GetParamSocSecur('SO_MAILCABINET','') ;
Chemin:=GetParamSocSecur('SO_REPERTOIREREVISION','') ;
if Chemin='' then MSG.Execute(2,Caption,'') ;
if ((Chemin<>'') and (Copy(Chemin,Length(Chemin),1)<>'\')) then Chemin:=Chemin+'\' ;
Masque:=GetParamSocSecur('SO_MASQUEREVISION','') ;
if Masque='' then MSG.Execute(3,Caption,'') ;
EdPath.Text:=Chemin+Masque ;
END ;

procedure TFRevDist.FormShow(Sender: TObject);
var Chemin,Masque,s : string;
begin
inherited;
PositionneDate ;
ChargeParamsCabinet ;
RenseigneRG ;
RempliCorps ;
FMailClick(nil) ;
end;

procedure TFRevDist.bExecClick(Sender: TObject) ;
begin
case Rg.ItemIndex of
   0 : Transactions(Envoyer,3) ;
   1 : Transactions(Recevoir,3) ;
   2 : AnnulerEnvoi ;
   end ;
end;

procedure TFRevDist.Recevoir ;
var Err     : TPGIErr ;
    Fichier : string ;
    hFile   : TextFile ;
begin
Fichier:=EdPath.Text ;
if (Trim(Fichier)='') or (Not FileExists(Fichier)) then BEGIN MSG.Execute(4,Caption,'') ; Exit ; END ;
//if Not ImportTOBS1(Fichier, True, false, false, Err) then
if Not IntegreFichierTOB(Fichier) then BEGIN MSG.Execute(5,Caption,Fichier+' : '+Err.Libelle) ; Exit ; END ;
if ((Not V_PGI.SAV) and (FileExists(Fichier)) then BEGIN AssignFile(hFile,Fichier) ; Erase(hFile) ; END ;
MSG.Execute(6,Caption,'') ;
DeleteFile(Fichier) ;
end;

Function TFRevDist.EnvoiPossible : boolean ;
Var Fichier : String ;
BEGIN
Result:=False ;
Fichier:=EdPath.Text ;
if Not V_PGI.ModeExpert then
   BEGIN
   if IsRevision(strToDate(E_DATECOMPTABLE.Text)) then BEGIN MSG.Execute(7,Caption,'') ; Exit ; END ;
   END ;
if Not NomFichierCorrect(Fichier) then BEGIN MSG.Execute(8,Caption,'') ; Exit ; END ;
if FileExists(Fichier) then
   BEGIN
   if MSG.Execute(9,Caption,'')<>mrYes then Exit ;
   DeleteFile(Fichier) ;
   END ;
Result:=True ;
END ;

procedure TFRevDist.Envoyer ;
var dFin : TDateTime ;
    q,qw : TQuery ;
    NumPaquet,CurrNum : integer ;
    Max,nb : integer ;
    Ajout  : Boolean ;
    st,Fichier,sujet : string ;
    Importance : integer ;
begin
if Not EnvoiPossible then Exit ;
SourisSablier;
Fichier:=EdPath.Text ;
if V_PGI.ModeExpert then
   BEGIN
   Qw:=OpenSQL('Select Max(E_PAQUETREVISION) from ECRITURE',True) ;
   if Not Qw.EOF then Max:=Qw.Fields[0].AsInteger else Max:=0 ;
   Ferme(Qw) ;
   Q:=OpenSQL('Select E_NUMEROPIECE, E_NUMLIGNE, E_IO, E_PAQUETREVISION from ECRITURE '+
              'Where E_IO="" OR E_ETATREVISION="M" '+
              'OR E_PAQUETREVISION='+IntToStr(max)+ ' OR E_PAQUETREVISION=0 '+
              'Order by E_NUMEROPIECE, E_NUMLIGNE',True) ;
   END else
   BEGIN
   dFin:=StrToDate(E_DATECOMPTABLE.Text) ;
   Q:=OpenSql('Select E_NUMEROPIECE, E_NUMLIGNE, E_IO, E_PAQUETREVISION from ECRITURE '+
              'Where (E_IO="" AND E_DATECOMPTABLE<="'+UsDateTime(dFin)+'") '+
              'OR E_ETATREVISION="M" '+
              'Order by E_NUMEROPIECE, E_NUMLIGNE',True) ;
   END ;
if Q.EOF then
   BEGIN
   Ferme(Q) ;
   if Not V_PGI.ModeExpert then MSG.Execute(10,Caption,E_DATECOMPTABLE.Text)
                           else MSG.Execute(11,Caption,'') ;
   SourisNormale ;
   Exit ;
   END ;
if Not InitExportT0B(Fichier) then BEGIN Ferme(Q) ; Exit ; END ;
if Expert then NumPaquet:=0 else
   BEGIN
   NumPaquet:=GetParamSocSecur('SO_COMPTEURREVISION',0) ; Inc(NumPaquet) ;
   SetParamSoc('SO_COMPTEURREVISION',NumPaquet) ;
   END ;
CurrNum:=-1 ; InitMove(100,'') ;
While Not Q.EOF do
   BEGIN
   MoveCur(False) ;
   if CurrNum<>Q.FindField('E_NUMEROPIECE').AsInteger then
      BEGIN
      Ajout:=((Q.FindField('E_IO').AsString='') and (V_PGI.ModeExpert)) or (Q.FindField('E_PAQUETREVISION').AsInteger=0) ;
      ExportPiece(Fichier,Q.FindField('E_NUMEROPIECE').AsInteger,Ajout,False,NumPaquet) ;
      END ;
   CurrNum:=Q.FindField('E_NUMEROPIECE').AsInteger ;
   Q.Next ;
   END ;
Ferme(Q) ; FiniMove ;
if FMail.Checked then
   BEGIN
   Sujet:=HDiv.Mess[11] ;
   if FHigh.Checked then Importance:=2 else Importance:=1 ;
   SendMail(Sujet,FEMail.Text,'',TStringList(FCorpsMail.Lines),EdPath.Text,Importance) ;
   END ;
SetParamSoc('SO_DATEDERNREVISION',NowH) ;
if FMail.Checked then MSG.Execute(12,Caption,'') else MSG.Execute(13,Caption,'') ;
SourisNormale ;
// JLD if fArrete.Checked then ClotureMensuelle (nil, StrToDate(E_DATECOMPTABLE.text) );
end;

procedure TFRevDist.ExportPerimetreLettrage ( TOBEcr : TOB ; Fichier : string ; Ajout : Boolean ; NumPaquet : Integer ) ;
var Gene,Auxi,Lett,stSql : string ;
    TOBE             : TOB ;
    i,NumLigne,NumeroPiece,CurrNum : integer ;
    Q : TQuery ;
begin
if TOBEcr.Detail.Count<=0 then Exit ;
NumeroPiece:=TOBEcr.Detail[0].GetValue('E_NUMEROPIECE') ;
for i:=0 to TOBECR.Detail.Count-1 do
    BEGIN
    TOBE:=TOBEcr.Detail[i] ;
    Auxi:=TOBE.GetValue('E_AUXILIAIRE') ; Gene:=TOBE.GetValue('E_GENERAL') ;
    Lett:=TOBE.GetValue('E_LETTRAGE') ;
   //Recherche du périmetre de lettrage
    if Lett<>'' then
       BEGIN
       NumLigne:=TOBE.GetValue('E_NUMLIGNE') ;
       stSql:='Select * from ECRITURE Where E_AUXILIAIRE="'+Auxi+'" AND E_GENERAL="'+Gene+'" AND E_LETTRAGE="'+Lett+'" AND E_ETATLETTRAGE<>"AL" AND E_ETATLETTRAGE<>"RI" ' ;
       //exclure les pièces repondant au critere d'envoi de révision en cours
       if V_PGI.ModeExpert then stSql:=stSql+ 'AND E_ETATREVISION<>"M" AND E_PAQUETREVISION<>0 AND E_PAQUETREVISION<>'+IntToStr(NumPaquet)+' '
        	           else stSql:=stSql+ 'AND E_ETATREVISION<>"M" ' ;
       stSql:=stSql+'Order by E_NUMEROPIECE, E_NUMLIGNE ' ;
       Q:=OpenSQL(stSql,True) ;
       if Not Q.EOF then
          BEGIN
          CurrNum:=0 ;
          While Not Q.EOF do
             BEGIN
             if CurrNum<>Q.FindField('E_NUMEROPIECE').AsInteger then
                BEGIN
                CurrNum:=Q.FindField('E_NUMEROPIECE').AsInteger ;
                if ((CurrNum=NumeroPiece) and (Q.FindField('E_NUMLIGNE').AsInteger<>NumLigne)) or
                    (CurrNum<>NumeroPiece) then ExportPiece(Fichier,CurrNum,False,True,NumPaquet) ;
                END ;
             Q.Next ;
             END ;
          END ;
       Ferme(Q) ;
       END ;
    END ;
end ;

procedure TFRevDist.ExportPiece ( Fichier : string ; NumeroPiece : Integer ; Ajout,SansPerimetre : boolean ; NumPaquet : integer ) ;
var TOBEcr : TOB ;
    i : integer;
begin
if Expert then E_FlagIO(NumeroPiece,'RC','') else E_FlagIO(NumeroPiece,'RE',intToStr(NumPaquet)) ;
TOBEcr:=TOB.Create('ECRITURES',Nil,-1) ;
// JLD E_Charge(NumeroPiece,TOBEcr) ;
if Not SansPerimetre then ExportPerimetreLettrage(TOBEcr,Fichier,Ajout,NumPaquet) ;
if Not Ajout then
   BEGIN
   TOBEcr.AddChampSup('ACTIONFICHE',False) ;
   TOBEcr.PutValue('ACTIONFICHE','MODIF') ;
   end;
//Annule l'état modifié et non de la base de l'émetteur
for i:=0 to TOBEcr.Detail.Count-1 do TOBEcr.Detail[i].PutValue('E_ETATREVISION','-') ;
TOBEcr.SaveToFile(Fichier,True,True,True) ;
TOBEcr.Free ;
end ;

procedure TFRevDist.edPathElipsisClick(Sender: TObject);
begin
if CM.Execute then EdPath.Text:=CM.Filename ;
end;

procedure TFRevDist.AnnulerEnvoi ;
var Paquet : string ;
    NumP : integer ;
begin
if MSG.Execute(14,Caption,'')<><>mrYes then Exit ;
if Not V_PGI.ModeExpert then
   BEGIN
   NumP:=GetParamSocSecur('SO_COMPTEURREVISION',0) ;
   if NumP>0 then
      BEGIN
      SourisSablier ;
      Paquet:=IntToStr(NumP) ; Dec(NumP) ; SetParamSoc('SO_COMPTEURREVISION',NumP) ;
      ExecuteSQL('Update ECRITURE Set E_IO="", E_PAQUETREVISION='+IntToStr(NumP)+' Where E_PAQUETREVISION='+Paquet) ;
      SourisNormale ;
      END ;
   END else
   BEGIN
   END ;
MSG.Execute(15,Caption,'') ;
end;

function TFRevDist.PreviousPage : TTabSheet ;
BEGIN
Result:=nil ;
if GetPage=Periode then Result:=Choix ;
if GetPage=Fichier then
   BEGIN
   Case Rg.ItemIndex of
      0 : if Expert then Result:=Choix else Result:=Periode ;
      1,2 : Result:=Choix ;
      end ;
   END ;
if GetPage=Mail then
   BEGIN
   Case Rg.ItemIndex of
      0,1 : Result:=Fichier ;
      2   : Result:=Choix ;
      end ;
   END ;
if GetPage=Resume then
   BEGIN
   Case Rg.ItemIndex of
      0 : Result:=Mail ;
      1 : Result:=Fichier ;
      2 : Result:=Choix ;
      end ;
   END; 
END ;

function TFRevDist.NextPage : TTabSheet ;
begin
Result:=nil ;
if GetPage=Choix then
   BEGIN
   Case Rg.ItemIndex of
      0 : if Expert then Result:=Fichier else Result:=Periode ;
      1 : Result:=Fichier ;
      2 : Result:=Resume ;
      end ;
   END ;
if GetPage=Periode then Result:=Fichier ;
if GetPage=Fichier then if Rg.ItemIndex=0 then Result:=Mail else Result:=Resume ;
if GetPage=Mail    then Result:=Resume ;
end ;

procedure TFRevDist.PChange(Sender: TObject);
begin
inherited;
if GetPage=Resume then
   BEGIN
   FLib2.Visible:=(rg.ItemIndex=0) and (not Expert) ; FVal2.Visible:=FLib2.Visible ;
   FLib3.Visible:=(rg.ItemIndex<2) ; FVal3.Visible:=FLib3.Visible ;
   FLib4.Visible:=(rg.ItemIndex=0) and (FMail.Checked) ; FVal4.Visible:=FLib4.Visible ;
   // Valeurs
   FVal1.Caption:=rg.Items[rg.ItemIndex] ;
   FVal2.Caption:=HDiv.Mess[12]+ ' '+E_DATECOMPTABLE.Text ;
   if FArrete.Checked then FVal2.Caption:=FVal2.Caption+HDiv.Mess[13] ;
   FVal3.Caption:=edPath.Text ;
   FVal4.Caption:=HDiv.Mess[14]+' '+FEMail.Text ;
   bExec.Visible:=True ; bExec.Default:=True ;
   END else
   BEGIN
   bExec.Visible:=False ; bExec.Default:=False ;
   END ;
end;

procedure TFRevDist.FMailClick(Sender: TObject);
begin
FEMail.Enabled:=FMail.Checked ;
FCorpsMail.Enabled:=FMail.Checked ;
FHigh.Enabled:=FMail.Checked ;
end;

end.
