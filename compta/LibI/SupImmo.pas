unit SupImmo;

// CA - 09/07/1999 - Etalonnage barre d'avancement

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, Db, DBTables, Hqry, HTB97, ComCtrls, HRichEdt,
  ExtCtrls, Grids, DBGrids, HDB, StdCtrls, ColMemo, Hctrls, hmsgbox, Mask,
  HEnt1,Hstatus,Paramdat, ImEnt, UiUtil, HPanel, HRichOLE, LookUp ;

  Procedure SuppressionImmo;
  procedure SupprimeFicheImmo(Code : string) ;

type
  TFSupImmo = class(TFMul)
    I_ETABLISSEMENT: THValComboBox;
    tI_ETABLISSEMENT: THLabel;
    HLabel1: THLabel;
    I_IMMO: TEdit;
    Nature: THLabel;
    I_NATUREIMMO: THValComboBox;
    I_QUALIFIMMO: THValComboBox;
    HLabel2: THLabel;
    HLabel3: THLabel;
    I_LIBELLE: TEdit;
    HLabel5: THLabel;
    HM: THMsgBox;
    I_COMPTEIMMO: THCritMaskEdit;
    Pzlibre: TTabSheet;
    Bevel5: TBevel;
    TI_TABLE0: TLabel;
    TI_TABLE1: TLabel;
    TI_TABLE2: TLabel;
    TI_TABLE3: TLabel;
    TI_TABLE4: TLabel;
    I_TABLE4: THCritMaskEdit;
    I_TABLE3: THCritMaskEdit;
    I_TABLE2: THCritMaskEdit;
    I_TABLE1: THCritMaskEdit;
    I_TABLE0: THCritMaskEdit;
    TI_TABLE5: TLabel;
    TI_TABLE6: TLabel;
    TI_TABLE7: TLabel;
    TI_TABLE8: TLabel;
    TI_TABLE9: TLabel;
    I_TABLE9: THCritMaskEdit;
    I_TABLE8: THCritMaskEdit;
    I_TABLE7: THCritMaskEdit;
    I_TABLE6: THCritMaskEdit;
    I_TABLE5: THCritMaskEdit;
    HLabel11: THLabel;
    I_LIEUGEO: THValComboBox;
    HLabel10: THLabel;
    I_COMPTELIE: THCritMaskEdit;
    HLabel4: THLabel;
    I_ORGANISMECB: THCritMaskEdit;
    HLabel7: THLabel;
    Label10: TLabel;
    Label12: TLabel;
    HLabel8: THLabel;
    HLabel9: THLabel;
    I_DATEAMORT: THCritMaskEdit;
    I_DATEAMORT_: THCritMaskEdit;
    I_DATEPIECEA_: THCritMaskEdit;
    I_DATEPIECEA: THCritMaskEdit;
    FExercice2: THValComboBox;
    procedure BOuvrirClick(Sender: TObject); override ;
    procedure FormShow(Sender: TObject);
    procedure KeyDate(Sender: TObject; var Key: Char);
    procedure FListeDblClick(Sender: TObject); override ;
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure SuppressionFicheImmo ;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure I_ORGANISMECBElipsisClick(Sender: TObject);
    procedure I_COMPTEIMMOElipsisClick(Sender: TObject);
  private
    FCodeImmo : string ;
//    procedure SuppressionFicheImmo ;
  public
    { Déclarations publiques }
  end;

implementation

uses imogen,Outils{$IFDEF SERIE1} {$ELSE},UtilPGI{$ENDIF};
{$R *.DFM}

Procedure SuppressionImmo;
var FMulImmo: TFSupImmo;
    PP : THPanel ;
begin
if imBlocage(['nrCloture','nrBatchImmo'],True,'nrBatchImmo') then exit;
PP:=FindInsidePanel ;
FMulImmo:=TFSupImmo.Create(Application) ;
FMulImmo.FNomFiltre:='MULVIMMOS' ; // a la place de MULM
FMulImmo.Q.Liste:='MULVIMMOS' ;  // a la place de MULM
if PP=Nil then
   BEGIN
  try
  FMulImmo.ShowModal ;
  finally
  FMulImmo.Free ;
    {$IFDEF SERIE1}
  Bloqueur('nrBatchImmo',False) ;
  {$ELSE}
  _Bloqueur('nrBatchImmo',False) ;
  {$ENDIF}

  end ;
   END else
   BEGIN
   InitInside(FMulImmo,PP) ;
   FMulImmo.Show ;
   END ;
end;

// Procedure de suppression d'une fiche Immo et de tous les enregistrements
// qui lui sont liés.
//procedure TFSupImmo.SuppressionFicheImmo ;
procedure SupprimeFicheImmo(Code : string) ;
var Q:TQuery;
{$IFDEF SERIE1}
{$ELSE}
    i:integer;
{$ENDIF}
begin
if ExecuteSQL ('DELETE FROM IMMO WHERE I_IMMO = "'+ Code + '"')<>1 then V_PGI.IoError:=oeSaisie ;
if V_PGI.IoError=oeOk then ExecuteSQL ('DELETE FROM IMMOAMOR WHERE IA_IMMO = "'+ Code + '"');
if V_PGI.IoError=oeOk then ExecuteSQL ('DELETE FROM IMMOLOG WHERE IL_IMMO = "'+ Code + '"');
if V_PGI.IoError=oeOk then ExecuteSQL ('DELETE FROM IMMOECHE WHERE IH_IMMO = "'+ Code + '"');
{$IFDEF SERIE1}
{$ELSE}
if V_PGI.IoError=oeOk then
begin
  for i:=1 to ImMaxAxe do
  begin
    ExecuteSQL('DELETE FROM VENTIL Where V_NATURE="IM'+IntToStr(i)+'" AND V_COMPTE="'+Code+'"');
  end;
end ;
{$ENDIF}
if V_PGI.IoError=oeOk then
begin
  {$IFDEF SERIE1}
  {$ELSE}
  // Suppression du lien avec l'écriture éventuelle rattachée à l'immobilisation
  ExecuteSQL ('UPDATE ECRITURE SET E_IMMO="", E_NUMEROIMMO=0, E_DATEMODIF="'+UsTime(NowH)+'" WHERE E_IMMO="'+Code+'"');
  {$ENDIF}
  // Mise à jour des liens éventuels avec l'immobilisation supprimée
  Q:=OpenSQL('SELECT I_IMMO,I_IMMOLIE,I_IMMOLIEGAR FROM IMMO WHERE I_IMMOLIEGAR="'+Code+'"'+
             'OR I_IMMOLIE="'+Code+'"',TRUE);
  if Not Q.EOF then
  begin
    Q.First;
    while not Q.Eof do
    begin
      if Q.FindField('I_IMMOLIE').AsString=Code then
        ExecuteSQL ('UPDATE IMMO SET I_IMMOLIE="" WHERE I_IMMO="'+Q.FindField('I_IMMO').AsString+'"');
      if Q.FindField('I_IMMOLIEGAR').AsString=Code then
        ExecuteSQL ('UPDATE IMMO SET I_IMMOLIEGAR="",I_DATEDEPOTGAR="'+USDateTime(iDate1900)+'",I_DEPOTGARANTIE=0 WHERE I_IMMO="'+Q.FindField('I_IMMO').AsString+'"');
      Q.Next;
    end;
  end ;
  Ferme(Q) ;
end ;
end;

procedure TFSupImmo.BOuvrirClick(Sender: TObject);
var
  i: integer;
begin
if Q.FindField('I_ETAT').AsString<>'OUV' then
begin
  if HM.execute(4,Caption,'')<>mrYes then exit ;
end;
if HM.execute(1,Caption,'')<>mrYes then exit ;
inherited;
if FListe.AllSelected then
   BEGIN
   InitMove(Q.RecordCount,'');
   Q.First;
   while Not Q.EOF do
     BEGIN
     MoveCur(False);
     FCodeImmo:=Q.FindField('I_IMMO').AsString ;
     if Transactions(SuppressionFicheImmo, 3)<>oeOk then
        BEGIN
        MessageAlerte(HM.Mess[0]) ;
        Break ;
        END ;
     Q.Next;
     END;
     FListe.AllSelected:=False;
   END
   ELSE
   BEGIN
   InitMove(FListe.NbSelected,'');
   for i:=0 to FListe.NbSelected-1 do
       BEGIN
       MoveCur(False);
       FListe.GotoLeBookmark(i);
       FCodeImmo:=Q.FindField('I_IMMO').AsString ;
       if Transactions(SuppressionFicheImmo, 3)<>oeOk then
          BEGIN
          MessageAlerte(HM.Mess[0]) ;
          Break ;
          END  else ImMarquerPublifi (True);   // CA le 02/02/2002
       END;
   FListe.ClearSelected;
   END;
FiniMove;
BChercheClick(Sender);  //EPZ 03/11/98
end;

procedure TFSupImmo.FormShow(Sender: TObject);
begin
{$IFDEF SERIE1}
I_ETABLISSEMENT.Visible := False;
tI_ETABLISSEMENT.Visible := False;
{$ENDIF}
I_DATEAMORT.text:=StDate1900 ; I_DATEAMORT_.text:=StDate2099 ;
I_DATEPIECEA.text:=StDate1900 ; I_DATEPIECEA_.text:=StDate2099 ;
inherited ;
end;

procedure TFSupImmo.KeyDate(Sender: TObject; var Key: Char);
begin
inherited ;
ParamDate(Self,Sender,Key) ;
end;

procedure TFSupImmo.FListeDblClick(Sender: TObject);
begin
if(Q.Eof)And(Q.Bof) then Exit ;
inherited ;
FicheImmobilisation(Q,Q.FindField('I_IMMO').AsString,taConsult,'') ;
end;

procedure TFSupImmo.HMTradBeforeTraduc(Sender: TObject);
begin
inherited ;
ImLibellesTableLibre(PzLibre,'TI_TABLE','I_TABLE','I') ;
end;

procedure TFSupImmo.SuppressionFicheImmo;
begin
  if not IsOpeEnCours (nil,FCodeImmo,false) then
  begin
    if  Q.FindField('I_CHANGECODE').AsString = '' then
      SupprimeFicheImmo(FCodeImmo)
    else  HM.Execute (3,Caption,'');
  end
  else HM.Execute(2,Caption,'');
end;


procedure TFSupImmo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
if isInside(Self) then
  BEGIN
  {$IFDEF SERIE1}
  Bloqueur('nrBatchImmo',False) ;
  {$ELSE}
  _Bloqueur('nrBatchImmo',False) ;
  {$ENDIF}
  Action:=caFree ;
  END ;
end;

procedure TFSupImmo.I_ORGANISMECBElipsisClick(Sender: TObject);
begin
  inherited;
  LookUpList (TControl (Sender),TraduireMemoire('Auxiliaire'),'TIERS','T_AUXILIAIRE','T_LIBELLE','T_NATUREAUXI="FOU"','T_AUXILIAIRE',True,2) ;
end;

procedure TFSupImmo.I_COMPTEIMMOElipsisClick(Sender: TObject);
var stWhere : string;
begin
  stWhere:='G_GENERAL<="'+VHImmo^.CpteLocSup+'" AND G_GENERAL>="'+VHImmo^.CpteImmoInf+'"' ;
  {$IFDEF SERIE1}
  LookupList(TControl(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,3)  ;
  {$ELSE}
  LookupList(TControl(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,1)  ;
  {$ENDIF}
end;

end.
