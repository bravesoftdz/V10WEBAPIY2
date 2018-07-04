unit TrtVerifContrat;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, Mask, ExtCtrls,MenuDispVerifContrat,MenuOLG,HEnt1, Buttons,Utob,
{$IFDEF EAGLCLIENT}

{$ELSE}
     dbTables,
{$ENDIF}
  AffEcheanceUtil,AffaireModifPiece,AffaireDuplic,FactUtil,DicoAf,Fe_Main;

type
  TFTRTCONTRAT = class(TForm)
    PParam: TPanel;
    PCompteRendu: TPanel;
    infoTemps: TLabel;
    bFin: TButton;
    bAnnuler: TButton;
    GBRepertoire: TGroupBox;
    RepLog: THCritMaskEdit;
    HLabel1: THLabel;
    MemoTrace: TMemo;
    Label1: TLabel;
    BStop: TBitBtn;
    GroupBox1: TGroupBox;
    HLabel4: THLabel;
    bAff: TCheckBox;
    HLabel2: THLabel;
    bPro: TCheckBox;
    HLabel3: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    bCalcul: TCheckBox;
    bGener: TCheckBox;
    NbTotAff: THLabel;
    NbTrtAff: THLabel;
    HLabel9: THLabel;
    HLabel10: THLabel;
    HLabel11: THLabel;
    HLabel12: THLabel;
    NbTotPro: THLabel;
    NbTrtPro: THLabel;
    HLabel7: THLabel;
    HLabel8: THLabel;
    HLabel13: THLabel;
    DebCal: THCritMaskEdit;
    FinCal: THCritMaskEdit;
    LibUneAff: TCheckBox;
    UneAff: THCritMaskEdit;
    AffDeb: THCritMaskEdit;
    LibBorne: THLabel;
    AffFin: THCritMaskEdit;
    HLabel14: THLabel;
    procedure bFinClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure UneAffElipsisClick(Sender: TObject);
    procedure LibUneAffClick(Sender: TObject);
  private
      FileLog  : Textfile;
      StopTrt    : Boolean;
      LigneMemo, CptAff, CptPro: Integer;

    { Déclarations privées }
    function  TraiteCtr : Boolean;
    function  LectureAffaires (StatutAff : String): Boolean;
    function  LectureUneAffaire : Boolean;
    Function  TraiteUneAffaire(TobAff : Tob) : Boolean;
    Procedure EnregTrace (AvecMemo : boolean; msg : string);
  end;



Procedure LanceTrtContrat ;

implementation

{$R *.DFM}

Procedure LanceTrtContrat ;
Var
    FTRTCONTRAT   : TFTRTCONTRAT ;
begin

FTRTCONTRAT :=TFTRTCONTRAT .Create(Application) ;

FTRTCONTRAT .ShowModal ;
FTRTCONTRAT .Free;
end;


procedure TFTRTCONTRAT.bFinClick(Sender: TObject);
begin
TraiteCtr;
PgiBoxAf ('Fin de traitement', 'traitement des contrats');
end;



procedure TFTRTCONTRAT.bAnnulerClick(Sender: TObject);
begin
close;
end;

//**** Fonction de recup GA => GMAO  *****
function TFTRTCONTRAT.TraiteCtr : Boolean;
Var Date : TDateTime;
    bOK : Boolean;
    Q : TQuery ;
    stBorne : string;
begin
BAnnuler.Enabled:=False; BFin.Enabled:=False;
StopTrt := False; LigneMemo := 0;
CptAff := 0; CptPro:=0;
AssignFile(FileLog ,RepLog.Text + 'TraiteCtr.txt');
if (FileExists(RepLog.Text+ 'TraiteCtr.txt') = TRUE) then Append (FileLog)
                                                     else Rewrite (FileLog);

EnregTrace (True, 'Début de traitement à : '+FormatDateTime('dd/mm/yyyy ttttt',NowH));

if libUneAff.Checked then
   LectureUneAffaire
else
   begin
   stBorne := '';
   if AffDeb.Text <> '' then stBorne := ' AND AFF_AFFAIRE>="' + AffDeb.Text + '"';
   if AffFin.Text <> '' then stBorne := stBorne + ' AND AFF_AFFAIRE<="' + AffFin.Text + '"';
   // init compteur pour suivi
   Q:= OpenSQL ('SELECT Count(AFF_AFFAIRE2) from AFFAIRE WHERE AFF_STATUTAFFAIRE="AFF"'+ stBorne, true);
   if Not Q.EOF then NbTotAff.Caption :=IntTostr(Q.Fields[0].AsInteger);
   Ferme(Q);
   Q:= OpenSQL ('SELECT Count(AFF_AFFAIRE2) from AFFAIRE WHERE AFF_STATUTAFFAIRE="PRO"'+ stBorne, true);
   if Not Q.EOF then NbTotPro.Caption :=IntTostr(Q.Fields[0].AsInteger);
   Ferme(Q);

   // **** traitement des affaires *****
   if bAff.checked then
      begin
      EnregTrace (True, 'Traitement des affaires - Début à ' + FormatDateTime('dd/mm/yyyy ttttt',NowH));
      bok := LectureAffaires('AFF');
      if not bOk then
         begin
         EnregTrace (True, 'Erreur traitement des affaires');
         end;
      end;

   // ******* traitement des propositions ******
   if bPro.checked then
      begin
      EnregTrace (True,'Traitement des Propositions - Début à ' + FormatDateTime('dd/mm/yyyy ttttt',NowH));
      bOk := LectureAffaires('PRO');
      if not bOk then
         begin
         EnregTrace (True, 'Erreur traitement des propositions');
         end;
      end;
   end; // fin multi affaires

CloseFile (FileLog);
BAnnuler.Enabled:=True; BFin.Enabled:=True;
end;

function TFTRTCONTRAT.LectureUneAffaire : Boolean;
Var TOBAff : TOB ;
    Q : TQuery;
    bOk : Boolean;
    stSQL : string;
begin
if UneAff.Text = '' then Exit;

TOBAff := Tob.Create ('AFFAIRE',nil,-1);
stSQL := 'SELECT AFF_AFFAIRE, AFF_TOTALHT, AFF_TOTALHTDEV, ';
stSQL := stSQL + 'AFF_TOTALTAXE, AFF_TOTALTAXEDEV, ';
stSQL := stSQL + 'AFF_TOTALTTC, AFF_TOTALTTCDEV, ';
stSQL := stSQL + 'AFF_MONTANTECHE,AFF_MONTANTECHEDEV,AFF_INTERVALGENER FROM AFFAIRE ';
stSQL := stSQL + 'WHERE AFF_AFFAIRE="'+ UneAff.Text +'"';

Q := nil;
try
Q := OPENSQL (stSQL,True);
if Not Q.EOF then TobAff.selectDB ('',Q) else exit;
finally
   Ferme(Q);
   end;

bOk := TraiteUneAffaire(TobAff);
Tobaff.UpdateDBTable (false);
if Not bOk then
   begin
   EnregTrace (True, 'pb updates Affaires');
   end;
TobAff.Free;
end;

function TFTRTCONTRAT.LectureAffaires (StatutAff : String) : Boolean;
Var maxCpt, Cpt , BorneDeb, BorneFin, i  : integer;
   Q : TQuery;
   TobAffs, Tobdet  : TOB;
   SQLAff, stBorneDeb, stBorneFin, stwhere : string;
   bOK : Boolean;
begin
result := False;  bOk := True;

// Découpage par tranche de compteur pour éviter des modules trop important
// Borne de début de compteur
if AffDeb.Text = '' then Cpt := 0 else Cpt := strToInt (Copy(AffDeb.Text,5,7));
// Borne de Fin
if AffFin.Text = '' then
   begin
   Q:= OpenSQL ('SELECT Max(AFF_AFFAIRE2) from AFFAIRE WHERE AFF_STATUTAFFAIRE="' +StatutAff + '"', true); // modif test AFFAIRE2
   if Not Q.EOF then
      begin
      if Q.Fields[0].AsString <> '' then
         maxCpt:=Q.Fields[0].AsInteger+1 ;
      end;
   Ferme(Q);
   end
else
   MaxCpt := strToInt (Copy(AffFin.Text,5,7));

if maxCpt = 0 then Exit;

EnregTrace (True, 'Premier Compteur :'+ IntTostr(Cpt));
EnregTrace (True, 'Dernier Compteur :'+ IntTostr(MaxCpt));

while Cpt < MaxCpt do
   begin
   Bornedeb := Cpt; BorneFin := BorneDeb + 100;  // modif test  + 1000...
   if BorneFin > MaxCpt then BorneFin := MaxCpt+1;
   stBorneDeb := Format('%*.*d',[7,7,BorneDeb]);   // modif test longueur partie 2 cegid ...
   stBorneFin := Format('%*.*d',[7,7,BorneFin]);
   stWhere := ' WHERE AFF_AFFAIRE2 >= "'+ stBorneDeb+ '" AND AFF_AFFAIRE2 < "'+ stBorneFin + '" AND AFF_STATUTAFFAIRE="' +StatutAff + '"';
   SQLAFF := 'SELECT AFF_AFFAIRE, AFF_TOTALHT, AFF_TOTALHTDEV,  AFF_TOTALTAXE, AFF_TOTALTAXEDEV,  AFF_TOTALTTC, AFF_TOTALTTCDEV,  AFF_MONTANTECHE,AFF_MONTANTECHEDEV,AFF_INTERVALGENER FROM AFFAIRE ' + stWhere + ' ORDER BY AFF_AFFAIRE2';   // modif Test Affaire 2

   if StopTrt then Exit;
   TobAffs := TOB.Create('',Nil,-1);

   Q := OpenSQL( SQLAff , True);
   if Not Q.EOF then
      begin
      EnregTrace (False, stWhere);
      Tobaffs.LoadDetailDB ('AFFAIRE','','',Q,True);
      end;
      Ferme(Q);

   for i := 0 To TobAffs.Detail.count-1 do
      begin
      if StatutAff = 'AFF' then Inc(CptAff) else Inc(CptPro);
      TobDet := TobAffs.Detail[i];
      bOk := TraiteUneAffaire(TobDet);
      end;
   if StatutAff = 'AFF' then NbTrtAff.Caption := IntTostr(CptAff) else NbTrtPro.Caption:=IntTostr(CptPro);


   // modif affaire pour réinsertion des totaux si recalcul
   if bCalcul.checked then
      bOk := TobAffs.UpdateDBTable (false);
   if Not bOk then
      begin
      EnregTrace (True, 'pb updates Affaires'+stWhere);
      end;

   TobAffs.Free; TobAffs := Nil;
   Cpt := BorneFin;
   end;
Result := True;
end;

Function TFTRTCONTRAT.TraiteUneAffaire (TobAff : Tob): Boolean;
Var Codeaff, RecupTotaux : string;
    DtDebCal, DtFinCal : TDateTime;
    bOk : Boolean;
    NumErr : integer;

begin
Result := False;
if StopTrt then Exit;
CodeAff := TobAff.GetValue('AFF_AFFAIRE');
EnregTrace (False,'Affaire :'+ CodeAff);


// **** Génération des échéances ****
if bGener.checked then
   begin
   DtDebcal := strToDate(DebCal.Text);  // modif test à voir sur quelle période on lance le traitement
   DtFinCal := strTodate(FinCal.Text);
   bOk := GenerEcheancesAffaire (CodeAff ,DtDebCal, DtFinCal );
   if Not bOk then
      begin
      EnregTrace (True,'pb Gener Ech. Affaire:'+ CodeAff); Exit;
      end;
   end;

// **** Recalcul de la piece ****
if bCalcul.checked then
   begin
   NumErr:= ModifieLaPieceAff(tmpRecalcul, CodeAff, Nil ,False,RecupTotaux);
   if numErr <> 0 then
      begin
      EnregTrace (True,'Pb Recalcul Affaire:'+ CodeAff + ' N°'+ intTostr(Numerr)); Exit;
      end;

   // **** Maj des totaux dans l'affaire ****
   RecalculTotAffaire (RecupTotaux ,TobAff);
      // Maj des contrevaleurs + pivot
      // Attention spécif CEGID en EURO comme monnaie pivot
   TobAff.PutValue('AFF_TOTALHT',TobAff.GetValue('AFF_TOTALHTDEV'));
   TobAff.PutValue('AFF_TOTALTAXE',TobAff.GetValue('AFF_TOTALTAXEDEV'));
   TobAff.PutValue('AFF_TOTALTTC',TobAff.GetValue('AFF_TOTALTTCDEV'));
   TobAff.PutValue('AFF_MONTANTECHE',TobAff.GetValue('AFF_MONTANTECHEDEV'));

// mise à jour du montant estimatif des échéances
   TobAff.PutValue('AFF_MONTANTECHEDEV',TobAff.GetValue('AFF_TOTALHTDEV')* TobAff.GetValue('AFF_INTERVALGENER'));
   TobAff.PutValue('AFF_MONTANTECHE',TobAff.GetValue('AFF_MONTANTECHEDEV'));
   end;

Result := True;
end;

Procedure TFTRTCONTRAT.EnregTrace (AvecMemo : boolean; msg : string);
begin
msg := IntToStr(LigneMemo)+' :'+Msg;
if AvecMemo then
   begin MemoTrace.Lines[LigneMemo]:= msg; Inc(LigneMemo); end;
writeln (FileLog, Msg);
end;


procedure TFTRTCONTRAT.BStopClick(Sender: TObject);
begin
StopTrt := True;
end;

procedure TFTRTCONTRAT.UneAffElipsisClick(Sender: TObject);
Var tmp,stArgument : string;
begin
stArgument := '';
// mcd 12/02/06 tmp:=AGLLanceFiche('AFF','AFFAIRERECH_MUL','','',StArgument);
tmp:=AFLanceFiche_AffaireRech('',StArgument);
if THCritMaskEdit(Sender).Tag = 0 then UneAff.Text := readtokenst(tmp);
if THCritMaskEdit(Sender).Tag = 1 then AffDeb.Text := readtokenst(tmp);
if THCritMaskEdit(Sender).Tag = 2 then AffFin.Text := readtokenst(tmp);
end;

procedure TFTRTCONTRAT.LibUneAffClick(Sender: TObject);
begin
UneAff.Visible := LibUneAFF.Checked;
end;

end.
