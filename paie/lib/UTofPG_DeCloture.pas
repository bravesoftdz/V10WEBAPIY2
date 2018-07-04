{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 25/06/2001
Modifié le ... :   /  /
Description .. : Décloture des paie
Mots clefs ... : PAIE;CLOTURE
*****************************************************************}
{
PT1 : 03/06/2002 V582 PH Utilisation HMtrad pour resize de la grille
PT2 : 03/06/2002 V582 PH Gestion historique des évènements
PT3 : 20/06/2004 V_60 PH FQ 11684 Pas de blocage monoposte

}

unit UTofPG_DECLOTURE;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids,
{$ENDIF}
      Grids,HCtrls,HEnt1,HMsgBox,HSysMenu,UTOF,UTOB,Vierge,P5Util,P5Def,AGLInit,
      EntPaie ;

const Tableau : Array [1..12] of String = (
 '01','02','03','04','05','06','07','08','09','10', '11','12');

Type
     TOF_PG_DECLOTURE = Class (TOF)
       private
       Grille : THGrid;
       VCbxLAnnee : THValComboBox;
       VCbxLMois : THValComboBox;
       IndiceClot : String;
       Trou : Integer;
       HMTrad    : THSystemMenu;
       procedure RecupPeriode;
       procedure RendCloture;
       function  RendMois (Indice : WORD) : WORD;
       procedure LanceDeCloture (DD, DF : TDateTime; MoisClos : String);
       procedure AnneeChange (Sender: TObject);
       procedure DeClotureEnter(Sender: TObject);
       public
       procedure OnClose ; override ;
       procedure OnArgument(Arguments : String ) ; override ;
     END ;

implementation

uses
    {$IFDEF STATDIR}
    DPMajPaieOutils,  // BTY 23/06/08 FQ 15526
    {$ENDIF}
    PgOutils2;


procedure TOF_PG_DECLOTURE.AnneeChange(Sender: TObject);
var   i, j, max : Integer;
begin
if Grille = NIL then exit;
RendCloture;
if IndiceClot= '' then IndiceClot := '------------';
RecupPeriode ;
max := 12;
if VH_Paie.PGDecalage = FALSE then i:=1 else i:=2;
if i = 2 then
 begin
 if IndiceClot[1] = 'X' then Grille.Cells[1,1] := TraduireMemoire ('Clos')
  else Grille.Cells[1,1] := TraduireMemoire ('A Cloturer');
 end;
for j:=i to max do
 begin
 if IndiceClot[j] = 'X' then begin Grille.Cells[1,j] := TraduireMemoire ('Clos'); Trou := i; end
  else  Grille.Cells[1,j] :=TraduireMemoire ('A Cloturer');
 end;
end;

procedure TOF_PG_DECLOTURE.DeClotureEnter(Sender: TObject);
var Min,Max,i,Mois  : WORD ;
    DD, DF  : TDateTime;
    St, MoisClos : String;
    Lig : Integer;
begin
if VCBXLMois = NIL then exit;
Mois := StrToInt(VCBXLMois.Value);
if VH_Paie.PGDecalage = TRUE then
 begin
 if Mois = 12 then Lig := 1
  else Lig := Mois + 1;
 end
 else Lig := Mois;
if Lig=0 then exit; // titre non accessible
if Grille.Cells [1, Lig] = TraduireMemoire ('A Cloturer') then
 begin
 PGIBox ('Vous ne pouvez pas déclôturer une période non close', 'Déclôture des Paies');
 exit; // Periode close on ne va pas la recloturer
 end
 else
 begin
 Max := Lig;
 Min := Max;
 for i:=1 to max do
  begin
  if IndiceClot [i] <> 'X' then begin Min := i; break; end;
  end;
 MoisClos:=IndiceClot;
 for i:=Min to max do
  begin MoisClos[i] := '-'; end;
 Mois := RendMois (Min);
 if VH_Paie.PGDecalage = FALSE then DD:=EncodeDate (StrToInt(vcbxLAnnee.Text),Mois,1) // Recup debut 1er mois A clore
  else
  begin
  if lig = 1 then DD:=EncodeDate (StrToInt(vcbxLAnnee.Text)- 1,Mois,1) // Annee - 1 car paie decalée
   else DD:=EncodeDate (StrToInt(vcbxLAnnee.Text) ,Mois,1);
  end;
 Mois := RendMois (Max);
 DF:=EncodeDate (StrToInt(vcbxLAnnee.Text),Mois,1);
 if (VH_Paie.PGDecalage = TRUE) AND (Mois = 12) then DF:=EncodeDate (StrToInt(vcbxLAnnee.Text) - 1,Mois,1);
 DF := FINDEMOIS (DF); // Recup dernier Mois à clore
 St := 'Voulez-vous déclôturer les paies du '+DateToStr(DD)+' au '+DateToStr(DF);
 if PGIAsk (St, 'Declôture des paies') = mrYes then
  begin
  LanceDeCloture (DD, DF, MoisClos);
  AnneeChange (NIL); // Pour reafficher les clotures effectuees
  end;
 end;
end;

procedure TOF_PG_DECLOTURE.LanceDeCloture(DD, DF: TDateTime; MoisClos : String);
var st,St1,S1,S2 : STring;
    Trace : TStringList;
    Err   : Boolean;
begin
// PT2 : 03/06/2002 V582 PH Gestion historique des évènements
Trace := TStringList.Create;
Err := FALSE;

st := 'UPDATE PAIEENCOURS SET PPU_TOPCLOTURE="-" WHERE PPU_DATEDEBUT>="'+UsDateTime (DD)+
      '" AND PPU_DATEFIN <="'+UsDateTime (DF)+'"';
try
 BeginTrans;
 ExecuteSql (st);
 st := 'UPDATE EXERSOCIAL SET PEX_CLOTURE="'+MoisClos+'" WHERE PEX_EXERCICE="'+vcbxLAnnee.value+'"';
 ExecuteSql (st);
 CommitTrans;
Except
 Rollback;
 PGIBox ('Une erreur est survenue lors de la déclôture', 'Déclôture des Paies');
 Err := TRUE;
END;

{$IFDEF STATDIR}
// BTY 23/06/08 FQ 15526
if (V_PGI.ModePCL='1') then DPMajDP_Paie (DD, DF, PGRendNoDossier,'DECLOTURE');
{$ENDIF}

// PT2 : 03/06/2002 V582 PH Gestion historique des évènements
S1:=DateToStr (DD);
S2:=DateToStr (DF);
St1 := '';
if Err then st1 := 'Erreur de ';
st := 'Déclôture des paies';
Trace.add (St1+St+' du '+S1+ ' au '+S2);
if Err then CreeJnalEvt ('001','004','Err', NIL,NIL,Trace)
 else CreeJnalEvt ('001','004','OK', NIL,NIL,Trace);
Trace.free;
// FIN PT2
end;

procedure TOF_PG_DECLOTURE.OnArgument(Arguments: String);
var    MoisE, AnneeE, ComboExer : String;
       BtnVal : TToolbarButton97;
       DebExer, FinExer : TDateTime;
begin
inherited ;
BtnVal := TToolbarButton97 (GetControl ('BValider'));
if BtnVal <> NIL then BtnVal.OnClick := DeClotureEnter;
// PT3 if NOT BlocageMonoPoste (TRUE) then exit;
Grille :=THGrid (GetControl('GRILLECLOTURE')) ;
VCbxLAnnee := THValComboBox (GetControl ('VCBXANNEE'));
VCbxLMois := THValComboBox (GetControl ('VCBXMOIS'));

if RendExerSocialEnCours (MoisE, AnneeE, ComboExer, DebExer, FinExer) = TRUE then
 begin
 if vcbxLAnnee <> NIL then
  begin
  Trou:=1;
  vcbxLAnnee.OnChange := AnneeChange;
  vcbxLAnnee.value:=ComboExer;
  if Trou < 12 then Grille.Row:=Trou+1 else Grille.Row := 1;
  end;
 end;
if VCbxLMois <> NIL then VCbxLMois.Value := MoisE;
//PT1 : 03/06/2002 V582 PH Utilisation HMtrad pour resize de la grille
HMTrad.ResizeGridColumns (Grille);

end;

procedure TOF_PG_DECLOTURE.OnClose;
begin
// PT3 DeblocageMonoPoste (TRUE);
end;
{ fonction de remplissage du libellé des mois dans la 1ere colonne en fonction du
  décalage de paie
}
procedure TOF_PG_DECLOTURE.RecupPeriode;
var i,j, max,val : Integer;
begin
if Grille = NIL then exit;
max := 12;
// On va remplir la 1ere colonne qui represente les mois de l'exercice social
if VH_Paie.PGDecalage = FALSE then begin i:=1; end
 else begin i:=2; end;
if i = 2 then
 begin
 Grille.Cells[0,1] := RechDOM ('PGMOIS', TABLEAU [12],FALSE);
 Grille.Cells[0,2] := RechDOM ('PGMOIS', TABLEAU [1],FALSE);
 i := i + 1;
 end;
for j:=i to max do
 begin
 if VH_Paie.PGDecalage = FALSE then val := j
  else val := j - 1;
 Grille.Cells[0,j] := RechDOM ('PGMOIS', TABLEAU [val],FALSE);
 end;
end;

function TOF_PG_DECLOTURE.RendMois(Indice: WORD): WORD;
begin
result:=Indice;
if VH_Paie.PGDecalage = TRUE then
 begin
 if Indice=1 then result:=12 else
  result := Indice - 1 ;
 end;
end;

{ Fonction qui recupere le tableau des mois cloturés
}
procedure TOF_PG_DECLOTURE.RendCloture;
var   Q : TQuery ;
begin
IndiceClot :='';
Q:=OpenSQL('SELECT PEX_CLOTURE FROM EXERSOCIAL WHERE PEX_EXERCICE="'+vcbxLAnnee.value+'"',TRUE) ;
if Not Q.EOF then
    BEGIN
    IndiceClot := Q.Fields[0].AsString ;
    END ;
Ferme (Q);
end;


Initialization
registerclasses([TOF_PG_DECLOTURE]);
end.
