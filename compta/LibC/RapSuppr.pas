{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 20/02/2003
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit Rapsuppr;

interface

uses
  Windows,  // PMinMaxInfo
  HSysMenu, Menus, Classes, Controls, Dialogs, Grids, Hctrls, StdCtrls, Buttons, ExtCtrls,
  Forms,    // TForm
  Messages, // TMessage, WM_GetMinMaxInfo
  SysUtils, // IntToStr, ExtractFilePath, ChangeFileExt, ExtractFileName, FormatDateTime, Now, FileExists, StrToDate, StrToInt, Trim
  HEnt1,    // String17, String3
  Ent1,     // VH, PopZoom, tzGeneral, tzTiers, AxeTotz
  Inifiles, // TIniFile
  HMsgBox,  // THMsgBox

//RR le 02042003 
//{$IFNDEF SANSCOMPTA}
  Hcompte,  // TGGeneral, TGTiers, THCpteEdit
//{$ENDIF}
//{$IFNDEF GCGC}
  //UTILEDT,
//{$ENDIF}

{$IFDEF EAGLCLIENT}

{$ELSE}
 {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFNDEF GCGC}
  RappType, // TInfoMvt
{$ENDIF}
  UTob, TntButtons, TntGrids
  ;

Procedure RapportdeSuppression(Laliste : TList ; OkSupp : Byte) ;
Procedure RapportdErreurMvt(Laliste : TList ; OkSupp : Byte ; Var Rep : Boolean ; EnBatch : boolean ) ;
Procedure RapportdErreurCptGen(Laliste : TList ; OkSupp : Byte ; Var Rep : Boolean) ;
procedure RapportErreurEnSerie( vListe : TList ; vTypeTraitement : Byte ; vNumErreur : Integer = 0 );

//RR {$IFNDEF SANSCOMPTA}
function  MajIEGeneral(Compte : String17 ; Lettrable,Ventilable : boolean) : boolean ;
function  MajIETiers(Compte : String17 ; Lettrable : boolean) : boolean ;
function  FirstModePaie(ModeRegle : String3) : String3 ;
//{$ENDIF}

type
  TFRapsuppr = class(TForm)
    Pbouton  : TPanel;
    FListe   : THGrid;
    Msg: THMsgBox;
    InfoExo: TComboBox;
    FListe2: THGrid;
    MsgQ: THMsgBox;
    TNbError: TLabel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    BImprimer: THBitBtn;
    Panel2: TPanel;
    BRepar: THBitBtn;
    BRecto: THBitBtn;
    HMTrad: THSystemMenu;
    BRechercher: THBitBtn;
    FindDialog: TFindDialog;
    TimerBatch: TTimer;
    BMenuZoom: THBitBtn;
    PopZ: TPopupMenu;
    ZGen: THBitBtn;
    ZAux: THBitBtn;
    ZAna: THBitBtn;
//{$IFNDEF SANSCOMPTA}
    Cache: THCpteEdit;
//{$ENDIF}
    procedure FormShow(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BRectoClick(Sender: TObject);
    procedure FListe2DblClick(Sender: TObject);
    procedure BReparClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BRechercherClick(Sender: TObject);
    procedure FindDialogFind(Sender: TObject);
    procedure TimerBatchTimer(Sender: TObject);
    procedure BMenuZoomClick(Sender: TObject);
//{$IFNDEF SANSCOMPTA}
    procedure ZGenClick(Sender: TObject);
    procedure ZAuxClick(Sender: TObject);
    procedure ZAnaClick(Sender: TObject);
//{$ENDIF}
    procedure BAideClick(Sender: TObject);
  private
   LaListe         : TList ;
   OkSupp          : Byte ;
   OkRep, FindFirst : Boolean ;
   WMinX,WMinY : Integer ;
   procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
   procedure PourVerif(T : Byte) ;
   procedure GereImportBatch ;
{$IFDEF EAGLCLIENT}
   function  GenererTOBEtat ( vGrille : THGrid ) : TOB ;
{$ENDIF}
//{$IFNDEF SANSCOMPTA}
   procedure ChercheLesComptes ;
//{$ENDIF}
  public
   EnBatch : boolean ;
  end;

type DelInfo = Class
               LeCod,LeLib,LeMess, LeMess2, LeMess3, LeMess4, Gen, Aux, Sect : String ;
               GenACreer,AuxACreer,SectACreer : Boolean ;
               Jal,Qualif,Axe : String ;
               NumP : Integer ;
               Chrono : Integer ;
               End ;


implementation

{$R *.DFM}

uses
{$IFNDEF IMP}
 {$IFNDEF CCMP}
  {$IFNDEF GCGC}
  {$IFNDEF CCADM}
   {$IFDEF EAGLCLIENT}

   {$ELSE}
     MvtError,  // ChercheError
   {$ENDIF}
   {$ENDIF CCADM}
  {$ENDIF}
 {$ENDIF}
{$ENDIF}

{$IFDEF MODENT1}
  CPProcGen,
  CPTypeCons,
  CPProcMetier,
{$ENDIF MODENT1}
{$IFNDEF GCGC}
   uLibWindows,
{$ENDIF}
{$IFDEF EAGLCLIENT}
   UtileAGL  // PrintDBGrid
{$ELSE}
   PrintDBG  // PrintDBGrid
{$ENDIF}
;

const cNomRapportCpt  = 'RapportCptCP.log';
      cNomRapportMvt  = 'RapportMvtCP.log';
      cNomRapportLet  = 'RapportLetCP.log';
      cNomRapportCor  = 'RapportCorCP.log';

procedure TFRapsuppr.FormShow(Sender: TObject) ;
Var i, NBL : Integer ;
    Quoi   : String ;
    DeuxGrid : Boolean ;
begin
BMenuZoom.Visible:=FALSE ;
if Not EnBatch then TimerBatch.Interval:=0 ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
FListe2.Visible:=False ; BRecto.Visible:=False ;
BRepar.Visible:=False ; OkRep:=False ;  BRechercher.Visible:=False ;
case OkSupp of
  0://Calcul du Scoring
    BEGIN
    Caption:=Msg.Mess[0] ;
    FListe.Options:=FListe.Options+[goColSizing];
    FListe.Cells[0,0]:=Msg.Mess[2] ; FListe.Cells[1,0]:=Msg.Mess[3] ; FListe.Cells[2,0]:=Msg.Mess[4] ;
    FListe.ColWidths[1]:=340 ; FListe.ColWidths[2]:=50 ;
    HelpContext:=7583500 ;
    END ;
  1: //Fiche par Défaut.
    BEGIN
    FListe.ColCount:=3 ;
    FListe.ColWidths[0]:=97 ; FListe.ColWidths[1]:=194 ; FListe.ColWidths[2]:=278 ;
    HelpContext:=0 ;
    END ;
  2:// 2 : Import/ Export.
    BEGIN
    Caption:=Msg.Mess[1] ;
    FListe.Cells[0,0]:=Msg.Mess[5] ; FListe.Cells[1,0]:=Msg.Mess[6] ;
    FListe.Cells[2,0]:=Msg.Mess[7] ;
    HelpContext:=0 ;
    FListe.ColWidths[1]:=120 ; FListe.ColWidths[2]:=350 ;
    END ;
  3://Fiche VerifCompta
    BEGIN
    FListe.ColCount:=5 ; TNbError.Visible:=true ;
    FListe2.Options:=FListe2.Options+[goColSizing];
    FListe.ColWidths[0]:=70 ; FListe.ColWidths[1]:=95 ; FListe.ColWidths[2]:=84 ;
    FListe.ColWidths[3]:=66 ; FListe.ColWidths[4]:=480 ;
    FListe.Cells[0,0]:=Msg.Mess[9] ; FListe.Cells[1,0]:=Msg.Mess[10] ; FListe.Cells[2,0]:=Msg.Mess[11] ;
    FListe.Cells[3,0]:=Msg.Mess[12] ; FListe.Cells[4,0]:=Msg.Mess[13] ;
    FListe.Titres.Clear ;
    FListe.Titres.Add(';'+'C'+';'+'S') ;
    FListe.Titres.Add(';'+'C'+';'+'S') ;
    FListe.Titres.Add(';'+'C'+';'+'S') ;
    FListe.Titres.Add(';'+'C'+';'+'D') ;
    FListe.Titres.Add(';'+'G'+';'+'S') ;
    FListe.SortEnabled:=true ;
    InfoExo.Visible:=False ; BRecto.Visible:=True ; BRecto.Hint:=Msg.Mess[8] ;
    BRechercher.Visible:=True ;
    HelpContext:=3210100 ;
    END ;
  4://Verif Compte
    BEGIN
    FListe.ColCount:=4 ; caption:=Msg.Mess[15] ;
    FListe.Cells[0,0]:=Msg.Mess[2] ; FListe.Cells[1,0]:=Msg.Mess[29] ;
    FListe.Cells[2,0]:=Msg.Mess[18] ; FListe.Cells[3,0]:=Msg.Mess[13] ;
    FListe.ColWidths[0]:=69 ; FListe.ColWidths[1]:=95 ;
    FListe.ColWidths[2]:=127 ; FListe.ColWidths[3]:=290 ;
    TNbError.Visible:=true ;
    FListe.ColAligns[0]:=TaCenter ; FListe.ColAligns[1]:=taLeftJustify ;
    BRepar.Visible:=V_PGI.SAV  ; BRechercher.Visible:=True ;
    HelpContext:=3210100 ;
    END ;
  5://Verif ANO
    BEGIN
    FListe.ColCount:=6 ; caption:=Msg.Mess[17] ; TNbError.Visible:=true ;
    FListe.ColAligns[0]:=TaCenter ; FListe.ColAligns[1]:=taLeftJustify ;
    FListe.Cells[0,0]:=Msg.Mess[2] ; FListe.Cells[1,0]:=Msg.Mess[18] ; FListe.Cells[2,0]:=Msg.Mess[19] ;
    FListe.Cells[3,0]:=Msg.Mess[20] ; FListe.Cells[4,0]:=Msg.Mess[21] ; FListe.Cells[5,0]:=Msg.Mess[22] ;
    FListe.ColWidths[0]:=94 ; FListe.ColWidths[1]:=130 ; FListe.ColWidths[2]:=86 ;
    FListe.ColWidths[3]:=86 ; FListe.ColWidths[4]:=86 ;  FListe.ColWidths[5]:=86 ;
    BRechercher.Visible:=True ;
    HelpContext:=0 ;
    END ;
  6:// 6 : Ctrl doublons.
    BEGIN
    Caption:=Msg.Mess[30] ;
    FListe.ColAligns[0]:=TaCenter ; FListe.ColAligns[1]:=TaCenter ;
    FListe.Cells[0,0]:=Msg.Mess[38] ; FListe.Cells[1,0]:=Msg.Mess[31] ;
    FListe.Cells[2,0]:=Msg.Mess[32] ;
    HelpContext:=0 ;
    FListe.ColWidths[1]:=120 ; FListe.ColWidths[2]:=350 ;
    END ;
  7://Fiche Verif Caisse
    BEGIN
    FListe.ColCount:=4 ; TNbError.Visible:=True ; caption:=Msg.Mess[34] ;
    FListe2.Options:=FListe2.Options+[goColSizing];
    FListe.ColWidths[0]:=111 ; FListe.ColWidths[1]:=199 ;
    FListe.ColWidths[2]:=116 ; FListe.ColWidths[3]:=116 ;
    FListe.Cells[0,0]:=Msg.Mess[2] ; FListe.Cells[1,0]:=Msg.Mess[18] ;
    FListe.Cells[2,0]:=Msg.Mess[19] ; FListe.Cells[3,0]:=Msg.Mess[20] ;
    FListe.Titres.Clear ;
    FListe.Titres.Add(';'+'C'+';'+'S') ;
    FListe.Titres.Add(';'+'C'+';'+'S') ;
    FListe.Titres.Add(';'+'D'+';'+'S') ;
    FListe.Titres.Add(';'+'D'+';'+'S') ;
    FListe.SortEnabled:=true ;
    InfoExo.Visible:=False ; BRecto.Visible:=False ;
    BRechercher.Visible:=True ;
    HelpContext:=0 ;
    END ;
  8: // Compte inexistant sur Contrôle en importation de mouvement
    BEGIN
    FListe.ColCount:=2 ; caption:=Msg.Mess[15] ; TNbError.Visible:=true ;
    FListe.ColAligns[0]:=TaCenter ; FListe.ColAligns[1]:=taLeftJustify ;
    FListe.Cells[0,0]:=Msg.Mess[2] ; FListe.Cells[1,0]:=Msg.Mess[9] ;
    FListe.ColWidths[0]:=120 ; FListe.ColWidths[1]:=420 ; BRepar.Visible:=FALSE  ;
    BRechercher.Visible:=True ;
    HelpContext:=0 ;
    END ;
  77://Verif Immos
    BEGIN
    FListe.ColCount:=3 ; caption:=Msg.Mess[37] ;
    FListe.Cells[0,0]:=Msg.Mess[29] ; FListe.Cells[1,0]:=Msg.Mess[18] ;
    FListe.Cells[2,0]:=Msg.Mess[13] ;
    FListe.ColWidths[0]:=105 ; FListe.ColWidths[1]:=125 ;
    FListe.ColWidths[2]:=320 ;
    TNbError.Visible:=true ;
    FListe.ColAligns[0]:=TaCenter ; FListe.ColAligns[1]:=taLeftJustify ;
    BRepar.Visible:=V_PGI.SAV  ; BRechercher.Visible:=True ;
    HelpContext:=0 ;
    END ;
  END ;

case OkSupp of
  0,1,2,6 : BEGIN
          for i:=0 to LaListe.Count-1 do
              BEGIN
              FListe.Cells[0,i+1]:=DelInfo(LaListe.Items[i]).LeCod ;
              FListe.Cells[1,i+1]:=DelInfo(LaListe.Items[i]).LeLib ;
              FListe.Cells[2,i+1]:=DelInfo(LaListe.Items[i]).LeMess ;
              FListe.RowCount:=FListe.RowCount+1 ;
              END ;
          if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
          END ;
 3      : BEGIN
          DeuxGrid:=False ;
          for i:=0 to LaListe.Count-1 do
              BEGIN
              Quoi:=DelInfo(LaListe.Items[i]).LeMess4 ;
              If (Quoi<>'L') then
                 BEGIN
                 BMenuZoom.Visible:=TRUE ;
                 FListe.Cells[0,i+1]:=DelInfo(LaListe.Items[i]).LeCod ;
                 FListe.Cells[1,i+1]:=DelInfo(LaListe.Items[i]).LeLib ;
                 FListe.Cells[2,i+1]:=DelInfo(LaListe.Items[i]).LeMess ;
                 FListe.Cells[3,i+1]:=DelInfo(LaListe.Items[i]).LeMess2 ;
                 FListe.Cells[4,i+1]:=DelInfo(LaListe.Items[i]).LeMess3 ;
//                 InfoExo.Items[i]:=Quoi ;
                 InfoExo.Items.Add(Quoi) ;
                 DeuxGrid:=True ;
                 FListe.RowCount:=FListe.RowCount+1 ;
                 END Else
              If Quoi='L' then
                 BEGIN
                 NBL:=FListe.RowCount-2 ;
                 If DeuxGrid then
                    BEGIN
                    FListe2.Cells[0,i-NBL+1]:=DelInfo(LaListe.Items[i]).LeCod ;
                    FListe2.Cells[1,i-NBL+1]:=DelInfo(LaListe.Items[i]).LeLib ;
                    FListe2.Cells[2,i-NBL+1]:=DelInfo(LaListe.Items[i]).LeMess ;
                    FListe2.Cells[3,i-NBL+1]:=DelInfo(LaListe.Items[i]).LeMess2 ;
                    FListe2.Cells[4,i-NBL+1]:=DelInfo(LaListe.Items[i]).LeMess3 ;
//                    InfoExo.Items[i]:=Quoi ;
                    InfoExo.Items.Add(Quoi) ;
                    FListe2.RowCount:=FListe2.RowCount+1 ;
                    END Else
                    BEGIN
                    FListe2.Cells[0,i+1]:=DelInfo(LaListe.Items[i]).LeCod ;
                    FListe2.Cells[1,i+1]:=DelInfo(LaListe.Items[i]).LeLib ;
                    FListe2.Cells[2,i+1]:=DelInfo(LaListe.Items[i]).LeMess ;
                    FListe2.Cells[3,i+1]:=DelInfo(LaListe.Items[i]).LeMess2 ;
                    FListe2.Cells[4,i+1]:=DelInfo(LaListe.Items[i]).LeMess3 ;
//                    InfoExo.Items[i]:=Quoi ;
                    InfoExo.Items.Add(Quoi) ;
                    FListe2.RowCount:=FListe2.RowCount+1 ;
                    END ;
                 END ;
              END ;
          { Avec 2 Grilles ? }
          if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
          if FListe2.RowCount>2 then FListe2.RowCount:=FListe2.RowCount-1 ;
          FListe.Visible:=(FListe.Cells[4,1]<>'') ;
          FListe2.Visible:=(FListe2.Cells[4,1]<>'') ;
          BRecto.Enabled:=(FListe.Cells[4,1]<>'') and (FListe2.Cells[4,1]<>'');
          If FListe.Visible then
             BEGIN
             caption:=Msg.Mess[8] ;
             TNbError.Caption:=IntToStr(FListe.RowCount-1)+' '+Msg.Mess[23];
             END else
          If FListe2.Visible then
             BEGIN
             caption:=Msg.Mess[14] ;
             TNbError.Caption:=IntToStr(FListe2.RowCount-1)+' '+Msg.Mess[23] ;
             END ;
          If BRecto.Enabled then FListe2.Visible:=False ;
          END ;
 4      : BEGIN
          for i:=0 to LaListe.Count-1 do
              BEGIN
              If DelInfo(LaListe.Items[i]).LeMess='G' then
                 BEGIN FListe.Cells[0,i+1]:=Msg.Mess[25] ; (*InfoExo.Items[i]:='G' ;*) InfoExo.Items.Add('G') ; END else
              If DelInfo(LaListe.Items[i]).LeMess='T' then
                 BEGIN FListe.Cells[0,i+1]:=Msg.Mess[26] ; (*InfoExo.Items[i]:='T' ;*) InfoExo.Items.Add('T') ; END else
              If DelInfo(LaListe.Items[i]).LeMess='S' then
                 BEGIN FListe.Cells[0,i+1]:=Msg.Mess[27] ; (*InfoExo.Items[i]:='S;'+DelInfo(LaListe.Items[i]).LeMess3 ;*) InfoExo.Items.Add('S;'+DelInfo(LaListe.Items[i]).LeMess3) ; END else
              If DelInfo(LaListe.Items[i]).LeMess='J' then
                 BEGIN FListe.Cells[0,i+1]:=Msg.Mess[28] ; (*InfoExo.Items[i]:='J' ;*) InfoExo.Items.Add('J') ; END ;

              if DelInfo(LaListe.Items[i]).LeMess='S' then
                 begin FListe.Cells[1,i+1]:='('+DelInfo(LaListe.Items[i]).LeMess3+') '+DelInfo(LaListe.Items[i]).LeCod ; end else
                 begin FListe.Cells[1,i+1]:=DelInfo(LaListe.Items[i]).LeCod ; end ;
              FListe.Cells[2,i+1]:=DelInfo(LaListe.Items[i]).LeLib ;
              FListe.Cells[3,i+1]:=DelInfo(LaListe.Items[i]).LeMess2 ;
              FListe.RowCount:=FListe.RowCount+1 ;
              END ;
          if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
          TNbError.Caption:=IntToStr(FListe.RowCount-1)+' '+Msg.Mess[23];
          END ;
 5      : BEGIN
          for i:=0 to LaListe.Count-1 do
              BEGIN
              FListe.Cells[0,i+1]:=DelInfo(LaListe.Items[i]).LeCod ;
              FListe.Cells[1,i+1]:=DelInfo(LaListe.Items[i]).LeLib ;
              FListe.Cells[2,i+1]:=DelInfo(LaListe.Items[i]).LeMess ;
              FListe.Cells[3,i+1]:=DelInfo(LaListe.Items[i]).LeMess2 ;
              FListe.Cells[4,i+1]:=DelInfo(LaListe.Items[i]).LeMess3 ;
              FListe.Cells[5,i+1]:=DelInfo(LaListe.Items[i]).LeMess4 ;
              FListe.RowCount:=FListe.RowCount+1 ;
              END ;
          if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
          TNbError.Caption:=IntToStr(FListe.RowCount-1)+' '+Msg.Mess[23];
          END ;
 7      : BEGIN
          for i:=0 to LaListe.Count-1 do
              BEGIN
              FListe.Cells[0,i+1]:=DelInfo(LaListe.Items[i]).LeCod ;
              FListe.Cells[1,i+1]:=DelInfo(LaListe.Items[i]).LeLib ;
              FListe.Cells[2,i+1]:=DelInfo(LaListe.Items[i]).LeMess ;
              FListe.Cells[3,i+1]:=DelInfo(LaListe.Items[i]).LeMess2 ;
              FListe.RowCount:=FListe.RowCount+1 ;
              END ;
          if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
          TNbError.Caption:=IntToStr(FListe.RowCount-1)+' '+Msg.Mess[23];
          END ;
 8      : BEGIN
          for i:=0 to LaListe.Count-1 do
              BEGIN
              FListe.Cells[0,i+1]:=DelInfo(LaListe.Items[i]).LeCod ;
              FListe.Cells[1,i+1]:=DelInfo(LaListe.Items[i]).LeLib ;
              FListe.RowCount:=FListe.RowCount+1 ;
              END ;
          if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
          TNbError.Caption:=IntToStr(FListe.RowCount-1)+' '+Msg.Mess[23];
          END ;
 77     : BEGIN
          for i:=0 to LaListe.Count-1 do
              BEGIN
              FListe.Cells[0,i+1]:=DelInfo(LaListe.Items[i]).LeCod ;
              FListe.Cells[1,i+1]:=DelInfo(LaListe.Items[i]).LeLib ;
              FListe.Cells[2,i+1]:=DelInfo(LaListe.Items[i]).LeMess2 ;
              FListe.RowCount:=FListe.RowCount+1 ;
              END ;
          if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
          TNbError.Caption:=IntToStr(FListe.RowCount-1)+' '+Msg.Mess[23];
          END ;
 end ;
UpdateCaption(Self) ;
if EnBatch then GereImportBatch ;
end;

procedure TFRapsuppr.GereImportBatch ;
var i,j : integer ;
    iniFile : TIniFile ;
    F : TextFile ;
    St,StPath,StName,Fic : String ;
BEGIN
If Not VH^.RecupSISCOPGI Then
  BEGIN
  StPath:=ExtractFilePath(Application.ExeName) ;
  IniFile:=TIniFile.Create(StPath+'IMPORT.INI');
  StName:=IniFile.ReadString('IMPORT','Fichier','');
  IniFile.Free ;
  Fic:=StPath+ChangeFileExt(ExtractFileName(StName),'.ERR') ;
  AssignFile(F,Fic) ;
  {$I-} ReWrite (F,Fic) ; {$I+}
  if IOResult<>0 then Exit ;
  Writeln(F,'Compte rendu d''importation du '+FormatDateTime('dddd dd mmmm yyyy "à" hh:nn',Now)) ;
  Writeln(F,'Fichier : '+StName) ; Writeln(F,'') ;
  for i:=1 to FListe.RowCount-1 do
    BEGIN
    St:='' ;
    for j:=0 to FListe.ColCount-1 do
      BEGIN
      if j<FListe.ColCount-1 then St:=St+Format_String(FListe.Cells[j,i]+' ',30)
                             else St:=St+Format_String(FListe.Cells[j,i],255) ;
      END ;
    Writeln(F,St) ;
    END ;
  CloseFile(F) ;
  END ;
TimerBatch.Enabled:=True ;
END ;

procedure TFRapsuppr.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFRapsuppr.BValiderClick(Sender: TObject);
begin
Case OkSupp of
  3 : BEGIN
      If FListe.Visible then PourVerif(1) else
      {JP 04/08/05 : FQ 15754 : FListe2 est utilisée pour le lettrage lorsque OkSupp est à 3 =>
                     donc PourVerif(3) et non PourVerif(2)}
      If FListe2.Visible then PourVerif(3) ;
      END ;
  4 : BEGIN
      PourVerif(4) ;
      END ;
  else Close ;
 end ;
end;

Procedure RapportdeSuppression(Laliste : TList ; OkSupp : byte) ;
var FRapsuppr: TFRapsuppr;
BEGIN
FRapsuppr:=TFRapsuppr.Create(Application) ;
 Try
  FRapsuppr.LaListe:=LaListe ;
  FRapsuppr.OkSupp:=OkSupp ;
  FRapsuppr.ShowModal ;
  Finally
  FRapsuppr.Free ;
 End ;
Screen.Cursor:=SynCrDefault ;
END ;

Procedure RapportdErreurMvt(Laliste : TList ; OkSupp : Byte ; Var Rep : Boolean ; EnBatch : boolean ) ;
var FRapsuppr: TFRapsuppr;
BEGIN
  // GCO - 26/04/2002
{$IFNDEF GCGC}
  if (CtxPcl in V_Pgi.PgiContexte) and (VH^.EnSerie) then
  begin
    RapportErreurEnSerie(LaListe, OkSupp);
    Exit;
  end;
{$ENDIF}
  // FIN GCO

FRapsuppr:=TFRapsuppr.Create(Application) ;
 Try
  FRapsuppr.LaListe:=LaListe ;
  FRapsuppr.OkSupp:=OkSupp ;
  FRapsuppr.EnBatch:=EnBatch ;
  FRapsuppr.BorderIcons := FRapsuppr.BorderIcons + [biMaximize];
  if OkSupp=3 then FRapsuppr.WIDTH:=630 ;
  FRapsuppr.ShowModal ;
  Rep:=FRapsuppr.OkRep ;
  Finally
  FRapsuppr.Free ;
 End ;
Screen.Cursor:=SynCrDefault ;
END ;

Procedure RapportdErreurCptGen(Laliste : TList ; OkSupp : Byte ; Var Rep : Boolean) ;
var FRapsuppr: TFRapsuppr;
BEGIN
FRapsuppr:=TFRapsuppr.Create(Application) ;
 Try
  FRapsuppr.LaListe:=LaListe ;
  FRapsuppr.OkSupp:=OkSupp ;
  FRapsuppr.ShowModal ;
  Rep:=FRapsuppr.OkRep ;
  Finally
  FRapsuppr.Free ;
 End ;
Screen.Cursor:=SynCrDefault ;
END ;

// GC - 13/02/2002
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/02/2002
Modifié le ... : 14/02/2002
Description .. : Fonction qui enregistre le contenu de vListe dans une tob
Suite ........ : qui est sauvegardée sur le disque
Suite ........ : vtypeTraitement vaut :  3 si Traitement des mouvements
Suite ........ :                         4 si Traitement des fiches
Mots clefs ... :
*****************************************************************}
procedure RapportErreurEnSerie( vListe : TList ; vTypeTraitement : Byte ; vNumErreur : Integer =0 );
{$IFNDEF GCGC}
var i : integer;
    lTobSave : Tob;
    lTobSave2 : Tob;
    lTobTemp : Tob;
    lSt : String;
{$ENDIF}
begin
{$IFNDEF GCGC}
  // Valeur de vTypeTraitement : 3 ----> Vérification sur les mouvements et le lettrage
  //                             4 ----> Vérification sur les comptes
  //                             5 ----> Erreur de Correction du programme

  lTobSave := Tob.Create('', nil, -1);
  lTobSave2 := Tob.Create('', nil, -1);
  try
    case vTypeTraitement of
      3 : begin
            if FileExists(GetWindowsTempPath + '\' + cNomRapportMvt) then TobLoadFromFile(GetWindowsTempPath + '\' + cNomRapportMvt, nil, lTobSave);
            if FileExists(GetWindowsTempPath + '\' + cNomRapportLet) then TobLoadFromFile(GetWindowsTempPath + '\' + cNomRapportLet, nil, lTobSave2);
          end;
      4 : if FileExists(GetWindowsTempPath + '\' + cNomRapportCpt) then TobLoadFromFile(GetWindowsTempPath + '\' + cNomRapportCpt, nil, lTobSave);
      5 : if FileExists(GetWindowsTempPath + '\' + cNomRapportCor) then TobLoadFromFile(GetWindowsTempPath + '\' + cNomRapportCor, nil, lTobSave);
    else
    end;

    if (vTypeTraitement = 3) or (vTypeTraitement = 4) then
    begin
      for i := 0 to vListe.Count - 1 do
      begin

        if vTypeTraitement = 3 then
        begin
          if DelInfo(vliste.Items[i]).LeMess4 = 'L' then
          begin
            lTobTemp := Tob.Create('', lTobSave2, -1);
            lTobTemp.AddChampSupValeur('Dossier', V_PGI.Nodossier);
            lTobTemp.AddChampSupValeur('Code', DelInfo(vListe.Items[i]).LeCod);
            lTobTemp.AddChampSupValeur('Auxiliaire', DelInfo(vListe.Items[i]).LeLib);
            lTobTemp.AddChampSupValeur('Général', DelInfo(vListe.Items[i]).LeMess);
            lTobTemp.AddChampSupValeur('Etat Lettrage', DelInfo(vListe.Items[i]).LeMess2);
            lTobTemp.AddChampSupValeur('Remarques', DelInfo(vListe.Items[i]).LeMess3);
          end
          else
          begin
            lTobTemp := Tob.Create('', lTobSave, -1);
            lTobTemp.AddChampSupValeur('Dossier', V_PGI.Nodossier);
            lTobTemp.AddChampSupValeur('Journal', DelInfo(vListe.Items[i]).LeCod);
            lTobTemp.AddChampSupValeur('Ref. interne', DelInfo(vListe.Items[i]).LeLib);
            lTobTemp.AddChampSupValeur('Pièce / Ligne', DelInfo(vListe.Items[i]).LeMess);
            lTobTemp.AddChampSupValeur('Date', DelInfo(vListe.Items[i]).LeMess2);
            lTobTemp.AddChampSupValeur('Remarques', DelInfo(vListe.Items[i]).LeMess3);
          end;
        end
        else
        begin
          if DelInfo(vListe.Items[i]).LeMess = 'G' then lSt := TraduireMemoire('Général');
          if DelInfo(vListe.Items[i]).LeMess = 'T' then lSt := TraduireMemoire('Tiers');
          if DelInfo(vListe.Items[i]).LeMess = 'S' then lSt := TraduireMemoire('Section');
          if DelInfo(vListe.Items[i]).LeMess = 'J' then lSt := TraduireMemoire('Journal');

          lTobTemp := Tob.Create('',lTobSave,-1);
          lTobTemp.AddChampSupValeur('Dossier', V_PGI.Nodossier);
          lTobTemp.AddChampSupValeur('Compte', lSt );
          if DelInfo(vListe.Items[i]).LeMess = 'S' then
            lTobTemp.AddChampSupValeur('Code', DelInfo(vListe.Items[i]).LeMess3 + ' ' + DelInfo(vListe.Items[i]).LeCod)
          else
            lTobTemp.AddChampSupValeur('Code', DelInfo(vListe.Items[i]).LeCod);
          lTobTemp.AddChampSupValeur('Libellé', DelInfo(vListe.Items[i]).LeLib);
          lTobTemp.AddChampSupValeur('Remarques', DelInfo(vListe.Items[i]).LeMess2);
        end;
      end;
    end //(vTypeTraitement = 3) or (vTypeTraitmement = 4)
    else
    begin
      lTobTemp := Tob.Create('',lTobSave,-1);
      lTobTemp.AddChampSupValeur('Dossier', V_PGI.Nodossier);
      lTobTemp.AddChampSupValeur('Type de traitement', 'Réparation des fichiers');
      lTobTemp.AddChampSupValeur('Détail du problème', 'Les erreurs ne peuvent pas être réparées "compte(s) mouvementé(s)');
    end;

    case vTypeTraitement of
      3 : begin // Mouvements et lettrage
            if lTobSave.Detail.Count  > 0 then lTobSave.SaveToFile(GetWindowsTempPath + '\' + cNomRapportMvt, False, True, True);
            if lTobSave2.Detail.Count > 0 then lTobSave2.SaveToFile(GetWindowsTempPath + '\' + cNomRapportLet, False, True, True);
          end;
      4 : lTobSave.SaveToFile(GetWindowsTempPath + '\' + cNomRapportCpt, False, True, True);
      5 : lTobsave.SaveToFile(GetWindowsTempPath + '\' + cNomRapportCor, False, True, True);
    else
    end;

  finally
    if lTobSave <> nil then lTobSave.Free;
    if lTobSave2 <> nil then lTobSave2.Free;
  end;
{$ENDIF}
end;
// GC- Fin

procedure TFRapsuppr.BImprimerClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
var lTOBEtat : TOB ;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  lTOBEtat := GenererTOBEtat(FListe) ;
  if OkSupp<>3 then
  	LanceEtatTOB( 'E' , 'RAS' , 'RAS' ,
                  lTOBEtat,
                  True, False, False, Nil, '', Caption, False )
  else
    begin
    if Fliste.Visible then
      LanceEtatTOB( 'E' , 'RAS' , 'RAS' ,
                    lTOBEtat,
                    True, False, False, Nil, '', Caption, False )
    else if Fliste2.Visible then
      LanceEtatTOB( 'E' , 'RAS' , 'RAS' ,
                    lTOBEtat,
                    True, False, False, Nil, '', Caption, False ) ;
    end;
  lTOBEtat.Free ;
{$ELSE}
  if OkSupp<>3 then begin
    PrintDBGrid(FListe,Nil,Caption,'') ;
    end
  else begin
    if Fliste.Visible  then
      PrintDBGrid(FListe,Nil,Caption,'')
    else if Fliste2.Visible then
      PrintDBGrid(FListe2,Nil,Caption,'') ;
  end;
{$ENDIF}
end;

procedure TFRapsuppr.FListeDblClick(Sender: TObject);
begin
//{$IFNDEF SANSCOMPTA}
case OkSupp of
  3 : if Pos('IMP',InfoExo.Items[FListe.Row-1])=0 then
      BEGIN
      If Pos(';',InfoExo.Items[FListe.Row-1])=0 Then PourVerif(1) Else PourVerif(2) ;
      END Else ChercheLesComptes ;
  4 : PourVerif(4) ;
  8 : ChercheLesComptes ;
  else Exit ;
 end ;
//{$ENDIF}
end;

procedure TFRapsuppr.PourVerif(T : Byte);
{$IFNDEF IMP}
{$IFNDEF CCMP}
{$IFNDEF GCGC}
Var PieLig, pie, LAxe,LEXO, PrefX, LesAxes, StTemp : String ;
    Mvt         : TInfoMvt ;
    I           : Byte ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
BEGIN
{$IFNDEF IMP}
{$IFNDEF CCMP}
{$IFNDEF GCGC}
{$IFNDEF CCADM}
Mvt:=TInfoMvt.Create ;
LAxe:=InfoExo.Items[FListe.Row-1] ;
LEXO:=ReadTokenSt(LAxe) ;
Case T of
   1 : BEGIN { Comptable Ecriture }
       Mvt.Journal:=FListe.Cells[0,FListe.Row] ;
       Mvt.Exercice:=LEXO ;
       Mvt.Datecomptable:=StrToDate(FListe.Cells[3,FListe.Row]) ;
       PieLig:=FListe.Cells[2,FListe.Row] ;
       while Pos('/',PieLig ) > 0 do PieLig[Pos('/', PieLig)] := ';' ;
       pie:=ReadTokenSt(PieLig) ;
       Mvt.NumeroPiece:=StrToInt(pie) ;
       Mvt.NumLigne:=StrToInt(PieLig) ;
       Mvt.INDICETABLE:=T ;
       END ;
   2 : BEGIN { Comptable Analytique }
       StTemp:=LAxe ; LesAxes:=LAxe ;
       ReadTokenSt(StTemp) ;
       For i:=1 to 5 do Mvt.AXEPLUS[i]:='' ;
       if (StTemp<>'') then
          BEGIN
          For i:=1 to 5 do if LesAxes<>'' then Mvt.AXEPLUS[i]:=ReadTokenSt(LesAxes) ;
          END Else Mvt.Axe:=LAxe ;

       Mvt.Journal:=FListe.Cells[0,FListe.Row] ;
       Mvt.Exercice:=LEXO ;
       Mvt.Datecomptable:=StrToDate(FListe.Cells[3,FListe.Row]) ;
       PieLig:=FListe.Cells[2,FListe.Row] ;
       while Pos('/',PieLig ) > 0 do PieLig[Pos('/', PieLig)] := ';' ;
       pie:=ReadTokenSt(PieLig) ;
       Mvt.NumeroPiece:=StrToInt(pie) ;
       Mvt.NumLigne:=StrToInt(PieLig) ;
       Mvt.INDICETABLE:=T ;
       END ;
   3 : BEGIN { Lettrage }
       Mvt.Lettrage:=FListe2.Cells[0,FListe2.Row] ;
       Mvt.Auxiliaire:=FListe2.Cells[1,FListe2.Row] ;
       Mvt.General:=FListe2.Cells[2,FListe2.Row] ;
       Mvt.EtatLettrage:=FListe2.Cells[3,FListe2.Row] ;
       Mvt.INDICETABLE:=T ;
       //OkZoom:=(Mvt.Lettrage<>'') and (Mvt.Auxiliaire<>'') and (Mvt.General<>'') ;
       END ;
   4 : BEGIN { Comptes }
       { En Fait Nature==Préfixe de la TABLE }
       if InfoExo.Items[FListe.Row-1]='G' then BEGIN Mvt.NATUREPIECE:='G' ; Mvt.General:=FListe.Cells[1,FListe.Row] ; END else
       if InfoExo.Items[FListe.Row-1]='T' then BEGIN Mvt.NATUREPIECE:='T' ; Mvt.Auxiliaire:=FListe.Cells[1,FListe.Row] ; END else
       if InfoExo.Items[FListe.Row-1]='J' then BEGIN Mvt.NATUREPIECE:='J' ; Mvt.Journal:=FListe.Cells[1,FListe.Row] ; END Else
          BEGIN
          LAxe:=InfoExo.Items[FListe.Row-1] ; PrefX:=ReadTokenSt(LAxe) ;
          Mvt.NATUREPIECE:='S' ; Mvt.Axe:=LAxe ;
          Mvt.Section:=Copy(FListe.Cells[1,FListe.Row],6,Length(FListe.Cells[1,FListe.Row]) ) ;
          END ;
       Mvt.INDICETABLE:=T ;
       END ;
     End ;
{$IFDEF EAGLCLIENT}
 // A FAIRE   Voir MvtError.PAS
{$ELSE}
  if V_PGI.SAV then ChercheError(Mvt);
{$ENDIF}
Mvt.Free ;
{$ENDIF CCADM}
{$ENDIF}
{$ENDIF}
{$ENDIF}
END ;

procedure TFRapsuppr.BRectoClick(Sender: TObject);
begin
If FListe.Visible then
   BEGIN
   FListe.Visible:=False ; FListe2.Visible:=True ;
   caption:=Msg.Mess[14] ; BRecto.Hint:=Msg.Mess[8] ;
   TNbError.Caption:=IntToStr(FListe2.RowCount-1)+' '+Msg.Mess[23] ;
   END Else
   BEGIN
   FListe2.Visible:=False ; FListe.Visible:=True ;
   caption:=Msg.Mess[8] ; BRecto.Hint:=Msg.Mess[14] ;
   TNbError.Caption:=IntToStr(FListe.RowCount-1)+' '+Msg.Mess[23] ;
   END ;
end;

procedure TFRapsuppr.FListe2DblClick(Sender: TObject);
begin
if OkSupp<>3 then exit ;
PourVerif(3) ;
end;

procedure TFRapsuppr.BReparClick(Sender: TObject);
begin
Case MsgQ.Execute(0,caption,'') of
  MrYes : begin OkRep:=True ; Close ; end ;
  MrNo  : Close ;
  End ;
end;

procedure TFRapsuppr.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Fliste.videPile(False) ; Fliste2.videPile(False) ;
end;

procedure TFRapsuppr.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFRapsuppr.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=240 ;
end;

procedure TFRapsuppr.BRechercherClick(Sender: TObject);
begin FindFirst:=True; FindDialog.Execute ; end;

procedure TFRapsuppr.FindDialogFind(Sender: TObject);
begin Rechercher(FListe,FindDialog, FindFirst); end;

procedure TFRapsuppr.TimerBatchTimer(Sender: TObject);
begin
TimerBatch.Enabled:=False ;
Close ;
end;

procedure TFRapsuppr.BMenuZoomClick(Sender: TObject);
begin
PopZoom(BMenuZoom,POPZ) ;
end;

function FirstModePaie(ModeRegle : String3) : String3 ;
var Q : TQuery ;
BEGIN
Result:='' ; if (ModeRegle='') then Exit ;
Q:=OpenSQL('SELECT MP_MODEPAIE FROM MODEPAIE ORDER BY MP_MODEPAIE',True,-1,'',true) ;
if not Q.Eof then Result:=Q.Fields[0].AsString ;
Ferme(Q) ;
Q:=OpenSQL('SELECT MR_MP1 FROM MODEREGL WHERE MR_MODEREGLE="'+ModeRegle+'"',True,-1,'',true) ;
if not Q.Eof then Result:=Q.Fields[0].asString ;
Ferme(Q) ;
END ;

//{$IFNDEF SANSCOMPTA}
function MajIEGeneral(Compte : String17 ; Lettrable,Ventilable : boolean) : boolean ;
var TGene : TGGeneral ;
BEGIN
Result:=False ;
TGene:=TGGeneral.Create(Compte) ;
if Lettrable then ExecuteSQL('UPDATE IMPECR SET IE_ECHE="X",IE_NUMECHE=1,IE_ETATLETTRAGE="AL",'+
                             'IE_LETTRAGE="",IE_DATEECHEANCE=IE_DATECOMPTABLE,IE_DATEPAQUETMIN=IE_DATECOMPTABLE,'+
                             'IE_DATEPAQUETMAX=IE_DATECOMPTABLE,IE_MODEPAIE="'+FirstModePaie(TGene.ModeRegle)+'"'+
                             ' WHERE IE_GENERAL="'+Compte+'" AND IE_ECHE="-"') ;
if Ventilable then ExecuteSQL('UPDATE IMPECR SET IE_ANA="X" WHERE IE_GENERAL="'+Compte+'" AND IE_ANA="-"') ;
TGene.free ;
END ;

function MajIETiers(Compte : String17 ; Lettrable : boolean) : boolean ;
var TTiers : TGTiers ;
BEGIN
Result:=False ;
TTiers:=TGTiers.Create(Compte) ;
if Lettrable then ExecuteSQL('UPDATE IMPECR SET IE_ECHE="X",IE_NUMECHE=1,IE_ETATLETTRAGE="AL",'+
                             'IE_LETTRAGE="",IE_DATEECHEANCE=IE_DATECOMPTABLE,IE_DATEPAQUETMIN=IE_DATECOMPTABLE,'+
                             'IE_DATEPAQUETMAX=IE_DATECOMPTABLE,IE_MODEPAIE="'+FirstModePaie(TTiers.ModeRegle)+'"'+
                             ' WHERE IE_AUXILIAIRE="'+Compte+'" AND IE_ECHE="-"') ;
TTiers.Free ;
END ;

procedure TFRapsuppr.ZGenClick(Sender: TObject);
Var St,Compte : String ;
    i,j,ColMaj : Integer ;
    Q : TQuery ;
    ACreer : Boolean ;
begin
St:='' ; If Fliste.Visible Then i:=FListe.Row Else Exit ;
ColMaj:=4 ; If OkSupp=8 Then ColMaj:=1 ;
Dec(i) ; If (i>=0) And (i<LaListe.Count) Then St:=DelInfo(LaListe.Items[i]).Gen ;
Cache.ZoomTable:=tzGeneral ; Cache.Text:=St ; ACreer:=FALSE ;
//SG6 18/01/05 FQ 15162
Compte := Cache.Text;
If Cache.Text<>'' Then
   BEGIN
   Q:=OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL="'+Cache.Text+'"',TRUE,-1,'',true) ;
   ACreer:=Q.Eof ; Ferme(Q) ;
   END ;

GChercheCompte(Cache,Nil);
If ACreer Then
   BEGIN
   //SG6 18/01/05 FQ 15162
   Q:=OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL="'+Compte+'"',TRUE,-1,'',true) ;
   If (Not Q.Eof) Then
      BEGIN
      MajIEGeneral(Cache.text,(Q.FindField('G_LETTRABLE').AsString='X'),(Q.FindField('G_VENTILABLE').AsString='X')) ;
      For j:=1 To Fliste.RowCount-1 Do
        BEGIN
        If (St=DelInfo(LaListe.Items[j-1]).Gen) And
           (DelInfo(LaListe.Items[j-1]).GenACreer) Then
           BEGIN
           FListe.Cells[ColMaj,j]:=MSg.Mess[35]+'('+FListe.Cells[ColMaj,j]+')' ;
           If OkSupp=8 Then DelInfo(LaListe.Items[j-1]).LeLib:=FListe.Cells[ColMaj,j] ;
           END ;
        END ;
      END ;
   Ferme(Q) ;
   END ;
end;

procedure TFRapsuppr.ZAuxClick(Sender: TObject);
Var St : String ;
    i,j,ColMaj : Integer ;
    Q : TQuery ;
    ACreer : Boolean ;
    Compte : string;
begin
St:='' ; If Fliste.Visible Then i:=FListe.Row Else Exit ;
ColMaj:=4 ; If OkSupp=8 Then ColMaj:=1 ;
Dec(i) ; If (i>=0) And (i<LaListe.Count) Then St:=DelInfo(LaListe.Items[i]).Aux ;
Cache.ZoomTable:=tzTiers ; Cache.Text:=St ; ACreer:=FALSE ;
//SG6 18/01/05 FQ 15162
Compte:= Cache.text;
If Cache.Text<>'' Then
   BEGIN
   Q:=OpenSQL('SELECT * FROM TIERS WHERE T_AUXILIAIRE="'+Cache.Text+'"',TRUE,-1,'',true) ;
   ACreer:=Q.Eof ; Ferme(Q) ;
   END ;
GChercheCompte(Cache,Nil);
If ACreer Then
   BEGIN
   //SG6 18/01/05 FQ 15162
   Q:=OpenSQL('SELECT * FROM TIERS WHERE T_AUXILIAIRE="'+Cache.Text+'"',TRUE,-1,'',true) ;
   If (Not Q.Eof) Then
      BEGIN
      MajIETiers(Cache.text,(Q.FindField('T_LETTRABLE').AsString='X')) ;
      For j:=1 To Fliste.RowCount-1 Do
        BEGIN
        If (St=DelInfo(LaListe.Items[j-1]).Aux) And
           (DelInfo(LaListe.Items[j-1]).AuxACreer) Then
           BEGIN
           FListe.Cells[ColMaj,j]:=MSg.Mess[35]+'('+FListe.Cells[ColMaj,j]+')' ;
           If OkSupp=8 Then DelInfo(LaListe.Items[j-1]).LeLib:=FListe.Cells[ColMaj,j] ;
           END ;
        END ;
      END ;
   Ferme(Q) ;
   END ;
end;

procedure TFRapsuppr.ZAnaClick(Sender: TObject);
Var St,Axe,Section : String ;
    i,j,ColMaj : Integer ;
    Q : TQuery ;
    ACreer : Boolean ;
begin
St:='' ; If Fliste.Visible Then i:=FListe.Row Else Exit ;
ColMaj:=4 ; If OkSupp=8 Then ColMaj:=1 ;
Dec(i) ;
If (i>=0) And (i<LaListe.Count) Then
  BEGIN
  St:=DelInfo(LaListe.Items[i]).Sect ;
  Axe:=DelInfo(LaListe.Items[i]).Axe ;
  END ;
If Trim(Axe)='' Then Exit ;
//SG6 18/01/05 FQ 15162
Section := Cache.Text;
Cache.ZoomTable:=AxeTotz(Axe) ; Cache.Text:=St ; ACreer:=FALSE ;
If Cache.Text<>'' Then
   BEGIN
   Q:=OpenSQL('SELECT * FROM SECTION WHERE S_SECTION="'+Cache.Text+'" AND S_AXE="'+Axe+'"',TRUE,-1,'',true) ;
   ACreer:=Q.Eof ; Ferme(Q) ;
   END ;
GChercheCompte(Cache,Nil) ;
If ACreer Then
   BEGIN
   //SG6 18/01/05 FQ 15162
   Q:=OpenSQL('SELECT * FROM SECTION WHERE S_SECTION="'+Section+'" AND S_AXE="'+Axe+'"',TRUE,-1,'',true) ;
   If (Not Q.Eof) Then
      For j:=1 To Fliste.RowCount-1 Do
        BEGIN
        If (St=DelInfo(LaListe.Items[j-1]).Sect) And
           (DelInfo(LaListe.Items[j-1]).SectACreer) And
           (DelInfo(LaListe.Items[j-1]).Axe=Axe) Then
           BEGIN
           FListe.Cells[ColMaj,j]:=MSg.Mess[35]+'('+FListe.Cells[ColMaj,j]+')' ;
           If OkSupp=8 Then DelInfo(LaListe.Items[j-1]).LeLib:=FListe.Cells[ColMaj,j] ;
           END ;
        END ;
   Ferme(Q) ;
   END ;
end;

procedure TFRapsuppr.ChercheLesComptes ;
Var i : Integer ;
    St : String ;
    XX : DelInfo ;
BEGIN
If Not FListe.Visible Then Exit ;
St:='' ; i:=FListe.Row-1 ;
If (i>=0) And (i<LaListe.Count) Then XX:=DelInfo(LaListe.Items[i]) Else Exit ;
If XX.GenACreer And (XX.Gen<>'') Then ZGenClick(Nil) Else
  If XX.AuxACreer And (XX.Aux<>'') Then ZAuxClick(Nil) Else
    If XX.SectACreer And (XX.Sect<>'') Then ZAnaClick(Nil) ;
END ;
//{$ENDIF}

procedure TFRapsuppr.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

{$IFDEF EAGLCLIENT}
function TFRapsuppr.GenererTOBEtat ( vGrille : THGrid ) : TOB;
var lInCpt    : Integer ;
    lTobLigne : TOB ;
begin
  Result := TOB.Create('GRILLE', nil, -1) ;
  for lInCpt := 1 to (vGrille.RowCount - 1) do
     begin
     lTobLigne := TOB.Create('RAPPORTSUPPR', Result , -1) ;
     lTobLigne.AddChampSup('RS_CODE',        False) ;
     lTobLigne.AddChampSup('RS_LIBELLE',     False) ;
     lTobLigne.AddChampSup('RS_COMMENT',     False) ;
     lTobLigne.PutValue( 'RS_CODE' ,         vGrille.Cells[ 0, lInCpt ] ) ;
     lTobLigne.PutValue( 'RS_LIBELLE' ,      vGrille.Cells[ 1, lInCpt ] ) ;
     lTobLigne.PutValue( 'RS_COMMENT' ,      vGrille.Cells[ 2, lInCpt ] ) ;
     end ;
end;
{$ENDIF}

end.
