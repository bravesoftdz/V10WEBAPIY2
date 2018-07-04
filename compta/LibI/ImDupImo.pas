{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 01/01/2004
Modifié le ... : 01/01/2004
Description .. : Option Duplication d'immo
Suite......... : FQ 12496 BTY 09/05 F10 doit valider la création de l'immo
Suite......... : FQ 18921 BTY 10/05 l'immo dupliquée ne doit pas contenir trace de DPI même si
Suite......... : l'immo d'origine en avait. Même chose si une prime d'équipement a été définie
Suite......... : BTY 11/06 Même chose si une subvention a été définie
Suite......... : BTY 19/11/07 Ne pas dupliquer paramètre agricole (passage du bien au régime réel)
Mots clefs ... : IMMO
*****************************************************************}

unit ImDupImo;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Mask, Hctrls, HTB97, ComCtrls, HRichEdt, HRichOLE, ExtCtrls,
  HPanel, hmsgbox,Db,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  HEnt1, HSysMenu, Utob, iment;

function ExecuteDuplication(CodeImmoOrig : string; var CodeImmoDest : string;AvecForm: boolean=true): TModalResult;

type
  TDupImmo = class(TForm)
    HPanel3: THPanel;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    GOrigine: TGroupBox;
    lCodeOrig: THLabel;
    lDateAchatOrig: THLabel;
    lQteAchatOrig: THLabel;
    lDesignationOrig: THLabel;
    lMontantHTOrig: THLabel;
    CodeOrig: TEdit;
    DateAchatOrig: TEdit;
    DesignationOrig: TEdit;
    QteAchatOrig: THNumEdit;
    MontantHTOrig: THNumEdit;
    GDestination: TGroupBox;
    lCodeDest: THLabel;
    lCodeDG: THLabel;
    lDesignationDest: THLabel;
    CodeDG: TEdit;
    CodeDest: TEdit;
    DesignationDest: TEdit;
    procedure BValiderClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CodeDestKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);

  private
    { Déclarations privées }
    T1,T2: TOB;
    GereDG: boolean ;
    function InitValeur(CodeOrigine: string): boolean ;
    function ValideFiche : boolean;
    function EnregDepotGarantie(CodeImmoLie: string): string ;
  public
    { Déclarations publiques }
  end;

implementation

uses Outils,ImPlan,IMMO_TOM{$IFDEF SERIE1} , Ut2Points, uterreur {$ENDIF}; //XMG 24/10/02

{$R *.DFM}

function ExecuteDuplication(CodeImmoOrig : string; var CodeImmoDest : string;AvecForm: boolean=true): TModalResult;
begin
result:=mrNone ;
with TDupImmo.Create(Application) do
  if InitValeur(CodeImmoOrig) then
    try
      ShowModal ;
    finally
      CodeImmoDest := CodeDest.Text;
      result:=ModalResult;
      Free ;
    end ;
end ;

procedure TDupImmo.FormCreate(Sender: TObject);
begin
  T1:=TOB.Create('IMMO',nil,-1) ;
  T2:=TOB.Create('IMMO',nil,-1) ;
  GereDG:=false ;
end;

procedure TDupImmo.FormDestroy(Sender: TObject);
begin
  T2.free ;
  T1.free ;
end;

function TDupImmo.InitValeur(CodeOrigine: string): boolean ;
begin
  with T1 do if SelectDB('"'+CodeOrigine+'"',nil) then
    begin
    // BTY fq 12496 09/05 Pouvoir intercepter la touche F10
    KeyPreview := True;
    CodeOrig.Text       :=GetValue('I_IMMO') ;
    DateAchatOrig.Text  :=GetValue('I_DATEPIECEA');
    MontantHTOrig.Text  :=FloatToStrF(GetValue('I_MONTANTHT'),ffFixed,20,V_PGI.OkDecV);
    QteAchatOrig.Text   :=FloatToStrF(GetValue('I_QUANTITE'), ffFixed,20,V_PGI.OkDecV);
    DesignationOrig.Text:=GetValue('I_LIBELLE');
    DesignationDest.Text:=DesignationOrig.Text;
    CodeDest.Text       :=NouveauCodeImmo;
    GereDG:=(GetValue('I_IMMOLIEGAR')<>GetValue('I_IMMO')) and T2.SelectDB('"'+GetValue('I_IMMOLIEGAR')+'"',nil) ;
    CodeDG.Enabled:=GereDG ; lCodeDG.Enabled:=GereDG ;
    if GereDG then CodeDG.Text:=Format('%.10d',[StrtoInt(CodeDest.Text)+1]) ;
    result:=true ;
    end
  else
    begin
    HM.Execute(9,Caption,'') ;
    result:=false ;
    end ;
end ;

procedure TDupImmo.BValiderClick(Sender: TObject);
var
//  infos: TInfoLog;
  TAmor, TLog, TEche, TUO : TOB;
  i : integer;
  AncienCode : string;
begin
  // Nouvelle méthode d'enregistrement en reprenant tout l'historique du plan et du log
  if ValideFiche then with T1 do
    begin
    AncienCode := T1.GetValue('I_IMMO');
    Putvalue('I_IMMO',CodeDest.Text) ;
    Putvalue('I_LIBELLE',DesignationDest.Text);
    Putvalue('I_CHANGECODE','');
    Putvalue('I_IMMOLIE','');
    // FQ 18921
    Putvalue('I_DPI','-'); // MVG - et pas vide
    Putvalue('I_SBVPRI',0);
    Putvalue('I_SBVPRIC',0);
    Putvalue('I_REPRISEUO',0);
    Putvalue('I_REPRISEUOCEDEE',0);
    Putvalue('I_CORVRCEDDE',0);
    Putvalue('I_SBVMT',0);
    Putvalue('I_SBVMTC',0);
    Putvalue('I_CORRECTIONVR',0);
    Putvalue('I_SBVDATE',iDate1900);
    Putvalue('I_CPTSBVB','');
    Putvalue('I_DPIEC','-');
    // BTY 19/11/07 Ne pas dupliquer paramètre agricole (passage du bien au régime réel)
    Putvalue('I_PFR','-');

    //
    if GereDG then Putvalue('I_IMMOLIEGAR',EnregDepotGarantie(GetValue('I_IMMOLIEGAR'))) ;
    TAmor := TOB.Create ('', T1, -1);
    TLog := TOB.Create ('', T1, -1);
    TEche := TOB.Create ('', T1, -1);
    TUO := TOB.Create('',T1,-1);
    try
      TAmor.LoadDetailDB ('IMMOAMOR','"'+AncienCode+'"','',nil,True);
      for i:=0 to TAmor.Detail.Count - 1 do
      begin
        TAmor.Detail[i].PutValue('IA_IMMO',T1.GetValue('I_IMMO'));
      end;
      TAmor.SetAllModifie (True);
      TLog.LoadDetailDB ('IMMOLOG','"'+AncienCode+'"','',nil,True);
      for i:=0 to TLog.Detail.Count - 1 do
      begin
        // on peut reprendre le log car pas d'opé en cours pour la duplication
        TLog.Detail[i].PutValue('IL_IMMO',T1.GetValue('I_IMMO'));
      end;
      TLog.SetAllModifie (True);
      TEche.LoadDetailDB ('IMMOECHE','"'+AncienCode+'"','',nil,True);
      for i:=0 to TEche.Detail.Count - 1 do
      begin
        TEche.Detail[i].PutValue('IH_IMMO',T1.GetValue('I_IMMO'));
      end;
      TEche.SetAllModifie (True);
      TUO.LoadDetailDB ('IMMOUO','"'+AncienCode+'"','',nil,True);
      for i:=0 to TUO.Detail.Count - 1 do
      begin
        TUO.Detail[i].PutValue('IUO_IMMO',T1.GetValue('I_IMMO'));
      end;
      TUO.SetAllModifie (True);
      SetAllModifie(true) ;
      InsertOrUpdateDB ;
      if GetValue('I_VENTILABLE')='X' then ImDupliqueVentil(CodeOrig.Text,CodeDest.Text) ;
    finally
      TAmor.Free;
      TLog.Free;
      TEche.Free;
      TUO.Free;
      // BTY fq 12496 09/05 Positionner modalresult si immo valide
      ModalResult:=mrYes;
    end;
  end
  else
    ModalResult:=mrNone;
end;

function TDupImmo.ValideFiche : boolean;
{$IFDEF SERIE1}
var Err : TPGIErr ; //XMG 24/10/02
{$ENDIF}
begin
  result:=false ;
  if CodeDest.Text='' then
    begin
    HM.Execute(6,Caption,'');
    FocusControl(CodeDest);
    end
  //XMG 24/10/02 début
{$IFDEF SERIE1}
  else if not Testecode(CodeDest.Text,Err) then
    Begin
    PGIBox(Err.Libelle,Caption) ;
    FocusControle(CodeDest) ;
    End
{$ENDIF}
//XMG 24/10/02 fin
  else if DesignationDest.Text='' then
    begin
    HM.Execute(7,Caption,'');
    FocusControl(DesignationDest);
    end
  else if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+CodeDest.Text+'" OR I_CHANGECODE="'+CodeDest.Text+'"') then
    begin
    HM.Execute(8,Caption,'');
    FocusControl(CodeDest);
    end
  else if GereDG and (ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+CodeDG.Text+'" OR I_CHANGECODE="'+CodeDG.Text+'"')) then
    begin
    HM.Execute(11,Caption,'');
    FocusControl(CodeDG);
    end
  else
    result:=true ;
end;

procedure TDupImmo.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

function TDupImmo.EnregDepotGarantie(CodeImmoLie: string): string ;
begin
with T2 do
  begin
  PutValue('I_IMMO',CodeDG.Text) ;
  PutValue('I_ABREGE',Copy(GetValue('I_IMMO'),1,17)) ;
  PutValue('I_LIBELLE',TraduireMemoire('Dépôt de garantie')+' '+T1.GetValue('I_LIBELLE')) ;
  PutValue('I_IMMOLIEGAR',T1.GetValue('I_IMMO')) ;
  PutValue('I_IMMOLIE','') ;
  SetAllModifie(true) ;
  InsertOrUpdateDB ;
  result:=GetValue('I_IMMO') ;
  end ;
end;


procedure TDupImmo.CodeDestKeyPress(Sender: TObject; var Key: Char);
begin
{$IFDEF SERIE1}
  //OkCar(Key,TRUE,#0) ; //XMG 24/10/02
{$ELSE}
  if (key='_') or (key='''') or (key='"') or (key='*') or (key='%') or (key='?') then
    key:=#0;
{$ENDIF}
end;

procedure TDupImmo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if (ModalResult<>mrYes) and  (HM.execute(10,Caption,'')<>mrYes) then Action:=caNone ;
end;

procedure TDupImmo.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var Action: TCloseAction;
begin
   // BTY fq 12496 09/05 bValiderClick insuffisant
   // if Key=VK_F10 then bValiderClick(nil);
   if Key=VK_F10 then
      begin
        bValiderClick(nil);
        If ModalResult = mrYes then OnClose(nil,Action);
      end;
end;

end.
