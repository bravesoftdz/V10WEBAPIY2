{$A-,H-}
unit IMPINT ;

interface

uses SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
     Forms, Dialogs, StdCtrls, Spin, DB, {$IFNDEF DBXPRESS}dbtables,
  hmsgbox, HSysMenu, Hqry, Buttons, ExtCtrls{$ELSE}uDbxDataSet{$ENDIF}, Hctrls, FileCtrl,GACCESS,
     ExtCtrls,Ent1, HStatus, ComCtrls, HEnt1, GDeclaDf, SaisUtil, Buttons, HCompte,
     hmsgbox, HSysMenu, LicUtil, Hqry ;


Procedure RecupINT ;

type
  TFImpINT = class(TForm)
    TNoteFam: THTable;
    TNoteMod: THTable;
    TChoixCod: THTable;
    TUser: THTable;
    TUserGrp: THTable;
    TDoc: THTable;
    Panel1: TPanel;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    Label1: TLabel;
    FRep: TEdit;
    BDir: TButton;
    FVider: TCheckBox;
    GroupBox1: TGroupBox;
    Image1: TImage;
    Image2: TImage;
    Image3: TImage;
    Image4: TImage;
    Image5: TImage;
    Image6: TImage;
    Image7: TImage;
    AImage1: TImage;
    AImage2: TImage;
    AImage3: TImage;
    AImage4: TImage;
    AImage5: TImage;
    AImage6: TImage;
    AImage7: TImage;
    OkGen: TCheckBox;
    OKFamille: TCheckBox;
    OkModele: TCheckBox;
    OkResp: TCheckBox;
    OkDoc: TCheckBox;
    OkModR: TCheckBox;
    OkRef: TCheckBox;
    Panel2: TPanel;
    BValider: TBitBtn;
    BAnnule: TBitBtn;
    BAide: TBitBtn;
    BSel: TBitBtn;
    BStop: TBitBtn;
    TMessage: TLabel;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    procedure RECUP(Sender: TObject);
    procedure BAnnuleClick(Sender: TObject);
    procedure BDirClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BStopClick(Sender: TObject);
  private { Private-déclarations }
    EnCours,Stopper   : boolean ;
    RepClient  : String ;
    RepModele  : String ;
    WordAPI    : Boolean ;
    procedure VideTable(NomTable,StSup : String) ;
    PROCEDURE LMess (st : String);
    procedure RecupSOCIETE ;
    procedure RecupMODELE ;
    procedure RecupFAMILLE ;
    procedure RecupRESP ;
    procedure RecupDOC ;
    FUNCTION SOCOK : Boolean ;
  public  { Public-déclarations }
  end;

implementation

{$R *.DFM}

{=============================================================================}
Function Bool(b:boolean) : String3 ;
begin
If b Then Bool:='X' Else Bool:='-' ;
end ;

{=============================================================================}
procedure TFImpINT.VideTable (NomTable,StSup : String) ;
Var St : String ;
BEGIN
if FVider.checked then
   BEGIN
   LMEss('Suppression en cours de '+NomTable) ;
   St:='DELETE FROM '+NomTable ;
   If STSup<>'' Then St:=St+' '+stSup ;
   ExecuteSQL(St) ;
   END ;
END ;

{=============================================================================}
PROCEDURE TFImpINT.LMess (st : String);

BEGIN
TMessage.Caption:=st ;
Application.ProcessMessages ;
END ;

{=============================================================================}
procedure TFImpINT.BAnnuleClick(Sender: TObject);
begin
//if Not Encours then Close else
//   BEGIN
//   Stopper:=TRUE ;
//   END ;
Close;
end;

{=============================================================================}
procedure TFImpINT.BDirClick(Sender: TObject);
var st : AnsiString ;
    i : integer ;
    Options: TSelectDirOpts ;
begin
If directoryExists(FRep.Text) then St:=FRep.text else st:='c:\' ;
i:=1 ; Options:=[] ;
if selectdirectory(st,Options,i) then FRep.text:=St ;
end;

{=============================================================================}
FUNCTION TFImpINT.SOCOK : Boolean ;
Var OkOk : Boolean ;
begin
okok:=FALSE ;
If directoryExists(FRep.Text) then
   begin
   OkOk:=FileExists(FRep.Text+'\SOC0\ACTION.DAT') ;
   if Not okok then ShowMessage(HM.Mess[0]) ;
   end else ShowMessage(HM.Mess[0]) ;
Result:=OKOK ;
end;

{=============================================================================}
procedure TFImpINT.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; 
Stopper:=FALSE ; EnCours:=FALSE ;
BStop.Enabled:=FALSE ;

end;

{=============================================================================}
procedure TFImpINT.RecupSOCIETE ;
Var Sorte : Byte ;
    PosFiche : LongInt ;
    Fiche : EnregINT ;
begin
Sorte:=9 ; PosFiche:=1 ;
Ouvre(Sorte,0,FALSE,spINT) ;
GetRec(VSAA^.Datff[Sorte],PosFiche,Fiche) ;
RepClient:='' ; RepModele:='' ; WordAPI:=FALSE ;
If Fiche.Status=0 Then
    BEGIN
    RepClient:=Fiche.RepClientS ;
    RepModele:=Fiche.RepModeleS ;
    WordAPI:=Fiche.WordAPIS ;
(*
    Fiche.RepGescomS  : String[50]   ;
    CreerClientS: Integer ;
    CreerNoteS  : Integer ;
    CreerFactS  : Integer ;
    CreerPassS  : Integer ;
    CreerTableS : Integer ;
    CreerCodeTS : Integer ;
    CreerCodeAS : Integer ;
    EncoursCalcS: Boolean ;
    NbLicS      : Integer ;*)
    END ;
GAccess.Ferme(Sorte,0,FALSE) ;
end ;

{=============================================================================}
procedure TFImpINT.RecupMODELE ;
Var Sorte : Byte ;
    PosFiche,Lg : LongInt ;
    Ext : String ;
    Fiche : EnregINT ;
begin
// VideTable('CHOIXCOD','WHERE CC_TYPE="APP"') ;
VideTable('NOTEMOD','') ;

Sorte:=5 ;
Ouvre(Sorte,0,FALSE,spINT) ; Lg:=FileLen(VSAA^.DatFF[Sorte])-1 ;
InitMove(Lg,'Recup des modèles de documents ') ;
BeginTrans ; TNoteMOD.open ; TChoixCod.Open ;
For PosFiche:=1 to lg do
   BEGIN
   If MoveCur(FALSE) Then ;
   GetRec(VSAA^.Datff[Sorte],PosFiche,Fiche) ;
   If Fiche.Status=0 Then
      BEGIN
      Ext:=Uppercase(Fiche.ExtensionY) ;
      if Not TChoixCod.FindKey(['APP',Ext]) then
         BEGIN
         TChoixCod.Insert ;
         TChoixCod.findfield('CC_TYPE').AsString:='APP' ;
         TChoixCod.findfield('CC_CODE').AsString:=Ext ;
         TChoixCod.findfield('CC_LIBELLE').AsString:=ExtractFileName(Fiche.CheminY) ;
         TChoixCod.findfield('CC_ABREGE').AsString:=ExtractFileName(Fiche.CheminY) ;
         TChoixCod.Post ;
         END ;
      TNoteMOD.Insert ;
      InitNew(TNoteMOD) ;
      TNoteMOD.findfield('NM_MODELE').asString:=FormatFloat('000',PosFiche) ;
      TNoteMOD.findfield('NM_LIBELLE').asstring:=Fiche.IntituleY ;
      TNoteMOD.findfield('NM_APPLICATION').asstring:=Ext ;
      if Trim(Fiche.ModeleDocY)<>'' then TNoteMOD.findfield('NM_NOMMODELE').asstring:=RepModele+Fiche.ModeleDocY ;
      TNoteMOD.findfield('NM_RESUME').asstring:=Bool(Fiche.ResumeY) ;
      if Fiche.AdresseY then
         BEGIN
         if WordAPI and (Ext='DOC') then TNoteMOD.findfield('NM_ADRESSE').asstring:='DET'
                                    else TNoteMOD.findfield('NM_ADRESSE').asstring:='GLO' ;
         END else TNoteMOD.findfield('NM_ADRESSE').asstring:='NUL' ;
      TNoteMOD.Post ;
      END ;
   END ;
TChoixCod.Close ; TNoteMOD.Close ;
CommitTrans ;
FiniMove ;
GAccess.Ferme(Sorte,0,FALSE) ;
end ;

{=============================================================================}
procedure TFImpINT.RecupFAMILLE ;
Var Sorte : Byte ;
    PosFiche,Lg : LongInt ;
    Fiche : EnregINT ;
begin
VideTable('NOTEFAM','') ;

Sorte:=6 ;
Ouvre(Sorte,0,FALSE,spINT) ; Lg:=FileLen(VSAA^.DatFF[Sorte])-1 ;
InitMove(Lg,'Recup des Familles de notes ') ;
BeginTrans ; TNoteFam.open ;
For PosFiche:=1 to lg do
   BEGIN
   If MoveCur(FALSE) Then ;
   GetRec(VSAA^.Datff[Sorte],PosFiche,Fiche) ;
   If Fiche.Status=0 Then
      BEGIN
      TNoteFam.Insert ;
      InitNew(TNoteFam) ;
      TNoteFam.findfield('NO_FAMILLE').asString:='6'+FormatFloat('00',PosFiche) ;
      TNoteFam.findfield('NO_SOUSFAMILLE').asstring:='-' ;
      TNoteFam.findfield('NO_LIBELLE').asstring:=Fiche.NomG ;
      TNoteFam.findfield('NO_REPERTOIRE').asstring:=RepClient+'\'+IntToStr(Fiche.NiveauG) ;
      TNoteFam.findfield('NO_NIVEAUACCES').asInteger:=Fiche.NiveauG ;
      TNoteFam.findfield('NO_UTILISATEUR').asstring:=V_PGI.User ;
      TNoteFam.Post ;
      END ;
   END ;
TNoteFAM.Close ;
CommitTrans ;
FiniMove ;
GAccess.Ferme(Sorte,0,FALSE) ;

//==================== Sous Famille ============================================
Sorte:=3 ;
Ouvre(Sorte,0,FALSE,spINT) ; Lg:=FileLen(VSAA^.DatFF[Sorte])-1 ;
InitMove(Lg,'Recup des Sous-Familles de notes ') ;
BeginTrans ; TNoteFam.open ;
For PosFiche:=1 to lg do
   BEGIN
   If MoveCur(FALSE) Then ;
   GetRec(VSAA^.Datff[Sorte],PosFiche,Fiche) ;
   If Fiche.Status=0 Then
      BEGIN
      TNoteFam.Insert ;
      InitNew(TNoteFam) ;
      TNoteFam.findfield('NO_FAMILLE').asString:='6'+FormatFloat('00',Fiche.CategorieN) ;
      TNoteFam.findfield('NO_SOUSFAMILLE').asstring:=FormatFloat('000',PosFiche) ;
      TNoteFam.findfield('NO_LIBELLE').asstring:=Fiche.NomN ;
      TNoteFam.findfield('NO_REPERTOIRE').asstring:=RepClient+'\'+IntToStr(Fiche.NiveauN) ;
      TNoteFam.findfield('NO_NIVEAUACCES').asInteger:=Fiche.NiveauN ;
      TNoteFam.findfield('NO_UTILISATEUR').asstring:=V_PGI.User ;
      TNoteFam.Post ;
      END ;
   END ;
TNoteFAM.Close ;
CommitTrans ;
FiniMove ;
GAccess.Ferme(Sorte,0,FALSE) ;
end ;

{=============================================================================}
procedure TFImpINT.RecupRESP ;
Var Sorte : Byte ;
    PosFiche,Lg : LongInt ;
    Groupe : String ;
    Fiche : EnregINT ;
begin
//VideTable('UTILISAT','') ;
TUserGrp.Open ;
if Not TUserGrp.Eof then
   BEGIN
   Groupe:=TUserGrp.FindField('UG_GROUPE').AsString ;
   End else Groupe:='' ;
TUserGrp.Close ;

Sorte:=10 ;
Ouvre(Sorte,0,FALSE,spINT) ; Lg:=FileLen(VSAA^.DatFF[Sorte])-1 ;
InitMove(Lg,'Recup des responsables ') ;
BeginTrans ; TUser.open ;
For PosFiche:=1 to lg do
   BEGIN
   If MoveCur(FALSE) Then ;
   GetRec(VSAA^.Datff[Sorte],PosFiche,Fiche) ;
   If Fiche.Status=0 Then
      BEGIN
      TUser.Insert ;
      InitNew(TUser) ;
      TUser.findfield('US_UTILISATEUR').asString:=FormatFloat('000',PosFiche) ;
      TUser.findfield('US_LIBELLE').asstring:=Fiche.MemoR ;
      TUser.findfield('US_ABREGE').asstring:=Fiche.NomR ;
      TUser.findfield('US_CONTROLEUR').asstring:='-' ;
      TUser.findfield('US_PASSWORD').asstring:=CryptageSt(Copy(Fiche.PassWordR,1,8)) ;
      TUser.findfield('US_PRESENT').asstring:='-' ;
//      TUser.findfield('US_DATECONNEXION').asstring:=Bool(Fiche.ResumeY) ;
      TUser.findfield('US_GROUPE').asstring:=Groupe ;
//      TUser.findfield('US_SUPERVISEUR').asstring:=Bool(Fiche.ResumeY) ;
      TUser.findfield('US_FONCTION').asstring:=Fiche.FonctionR ;
      TUser.Post ;
      END ;
   END ;
TUser.Close ;
CommitTrans ;
FiniMove ;
GAccess.Ferme(Sorte,0,FALSE) ;
end ;

{=============================================================================}
procedure TFImpINT.RecupDOC ;
Var Sorte : Byte ;
    PosFiche,Lg : LongInt ;
    Fiche : EnregINT ;
begin
VideTable('DOCUMENT','') ;
Sorte:=4 ;
Ouvre(Sorte,0,FALSE,spINT) ; Lg:=FileLen(VSAA^.DatFF[Sorte])-1 ;
InitMove(Lg,'Recup des responsables ') ;
BeginTrans ; TDoc.open ;
For PosFiche:=1 to lg do
   BEGIN
   If MoveCur(FALSE) Then ;
   GetRec(VSAA^.Datff[Sorte],PosFiche,Fiche) ;
   If Fiche.Status=0 Then
      BEGIN
      TDoc.Insert ;
(*                      4:(REFCLIT      : Longint    ;
                         DateT        : TDateTime  ;
                         NumeroT      : Integer    ;
                         MemoT        : String     ;
                         CommercialT  : Byte       ;
                         TypeActionT  : Integer    ;
                         NomFichierT  : String[12] ;
                         RefAFFT      : Longint    ;
                         NiveauT      : Integer    ;
                         NaTureT      : Byte        ;
                         ArappelerT   : TdateTime  ;
                         MotifT       : String[20] ;*)
      InitNew(TDoc) ;
      TDoc.findfield('DC_AUXILIAIRE').asString:=FormatFloat('000',PosFiche) ;
      TDoc.findfield('DC_FAMILLE').asstring:=Fiche.MemoR ;
      TDoc.findfield('DC_SOUSFAMILLE').asstring:=Fiche.NomR ;
      TDoc.findfield('DC_NUMERODOC').asstring:='-' ;
      TDoc.findfield('DC_LIBELLE').asstring:=Fiche.PassWordR ;
      TDoc.findfield('DC_NOMFICHIER').asstring:='-' ;
      TDoc.findfield('DC_APPLICATION').asstring:='-' ;
      TDoc.findfield('DC_MODELE').asstring:='-' ;
      TDoc.findfield('DC_NIVEAUACCES').asstring:='-' ;
      TDoc.findfield('DC_DATEMODIF').asstring:='-' ;
      TDoc.findfield('DC_DATECREATION').asstring:='-' ;
      TDoc.findfield('DC_UTILISATEUR').asstring:='-' ;
      TDoc.Post ;
      END ;
   END ;
TDoc.Close ;
CommitTrans ;
FiniMove ;
GAccess.Ferme(Sorte,0,FALSE) ;
end ;

{=============================================================================}
procedure TFImpINT.RECUP(Sender: TObject);
begin
if Not SocOK then exit ;
EnCours:=TRUE ;
SocPath:=Trim(FRep.text)+'\' ;
TMessage.Visible:=TRUE ;
BStop.Enabled:=TRUE ; BAnnule.Enabled:=FALSE ; BSel.Enabled:=FALSE ;
Image1.Visible:=TRUE ;  AImage1.Visible:=FALSE ;

RecupSOCIETE ;
If OkModele.Checked  Then BEGIN RecupMODELE   ; Image1.Visible:=TRUE ;  AImage1.Visible:=FALSE ; END ;
if OkFamille.Checked Then BEGIN RecupFAMILLE  ; Image2.Visible:=TRUE ;  AImage2.Visible:=FALSE ; END ;
if OkResp.Checked    Then BEGIN RecupRESP     ; Image3.Visible:=TRUE ;  AImage3.Visible:=FALSE ; END ;
if OkDoc.Checked     Then BEGIN RecupDOC      ; Image4.Visible:=TRUE ;  AImage4.Visible:=FALSE ; END ;
TMessage.Visible:=FALSE ;
EnCours:=FALSE ; Stopper:=FALSE ;
ChargerLesTables ;
BStop.Enabled:=FALSE ; BAnnule.Enabled:=TRUE ; BSel.Enabled:=TRUE ;
end;

Procedure RecupINT ;
Var FImpINT : TFImpINT ;
begin
if Not BlocageMonoPoste(False) then Exit ;
FImpINT:=TFImpINT.Create(Application) ;
  Try
   FImpINT.ShowModal ;
  Finally
   FImpINT.Free ;
   DeblocageMonoPoste(False) ;
  end ;
Screen.Cursor:=crDefault ;
end ;


procedure TFImpINT.BStopClick(Sender: TObject);
begin
Stopper:=TRUE ;
end;

end.
{$A+,H+}

