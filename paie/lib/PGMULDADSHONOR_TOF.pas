{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : PGMULDADSHONOR ()
Mots clefs ... : TOF;PGMULDADSHONOR
*****************************************************************}
{
PT1   : 20/02/2002 JL V571 Correction requête pour Oracle
PT2   : 26/04/2002 VG V571 Par défaut, année précédente jusqu'au mois de
                           septembre. A partir du mois d'octobre, année en cours
                           (pour tests)
PT3   : 15/07/2002 VG V585 Ajout traitement de duplication des honoraires
PT4   : 02/10/2002 VG V585 Duplication des honoraires : Correction pour ne pas
                           permettre la création d'honoraires avec date au
                           30/12/1899 - FQ N° 10260
PT5   : 05/11/2002 VG V585 Gestion du journal des evenements
PT6-1 : 29/01/2003 VG V591 Affichage d'un message si l'exercice par défaut
                           n'existe pas - FQ N°10469
PT6-2 : 12/02/2003 VG V_42 FQ N°10469 bis
PT7   : 20/02/2003 VG V_42 Cas du double-click sur une liste vide
PT8   : 26/02/2004 VG V_50 Traitement de suppression d'un honoraire - FQ N°11074
PT9   : 12/05/2004 VG V_50 Ajout du traitement lié aux raccourcis clavier
PT10  : 26/01/2005 VG V_60 MAJ de la liste après modification du nom
                           FQ N°11821
PT11  : 07/10/2005 VG V_60 Suite à modification sur les mul, modification du
                           chargement des listes
PT12-1: 16/10/2006 VG V_70 Moulinette de modification du champ PDS_SALARIE pour
                           les honoraires
PT12-2: 16/10/2006 VG V_70 Suppression du fichier de contrôle - mise en table
                           des erreurs
PT13  : 31/10/2006 VG V_70 Correction moulinette de modification du champ
                           PDS_SALARIE - FQ N°13637
PT14  : 21/11/2006 VG V_70 Permettre de déclarer des honoraires en DADS-U
                           complémentaire - FQ N°13613
PT15  : 24/11/2006 VG V_70 Correction moulinette de modification du champ
                           PDS_SALARIE - FQ N°13637
PT16  : 22/02/2007 VG V_72 Duplication des honoraires : mauvaise alimentation du
                           champ EXERCICEDADS
PT17  : 28/11/2007 VG V_80 Gestion du champ ET_FICTIF - FQ N°13925
}
unit PGMULDADSHONOR_TOF;

interface

uses Controls,
     Classes,
     sysutils,
     AGLInit,
     ed_tools,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     PgOutils2,
     hqry,
     PgDADSCommun,
     HStatus,
     HTB97,
     UTOB,
{$IFNDEF DADSUSEULE}
     P5Def,
{$ENDIF}
     Windows,
     PgDADSOutils,
{$IFDEF EAGLCLIENT}
     UtileAGL,
     MaineAgl,
     emul,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     HDB,
     FE_Main,
{$ENDIF}
     StdCtrls,
     EntPaie;
type
  TOF_PGMULDADSHONOR = class(TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    WW: THEdit; // Clause XX_WHERE
    THAnnee: THValCombobox;
    param: string;
    Q_Mul: THQuery;
    ComboAnnee2: THValComboBox;
{$IFNDEF EAGLCLIENT}
    Liste: THDBGrid;
{$ELSE}
    Liste: THGrid;
{$ENDIF}
    BCherche, Delete : TToolbarButton97;
    Compl : TCheckBox;

    procedure ActiveWhere (Sender: TObject);
    procedure GrilleDblClick (Sender: TObject);
    procedure GrilleClick (Sender: TObject);
    procedure Duplication(Sender: TObject);
    procedure Duplique_un(var OrdreDest:integer);
    procedure DeleteClick (Sender: TObject);
    procedure Delete_un();
    procedure NewHonor(Sender: TObject);
    procedure Annee2Change(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ComplClick (Sender: TObject);
    procedure MAJTypeD();
  end;

implementation

procedure TOF_PGMULDADSHONOR.OnNew;
begin
  inherited;
end;

procedure TOF_PGMULDADSHONOR.OnDelete;
begin
  inherited;
end;

procedure TOF_PGMULDADSHONOR.OnUpdate;
begin
Inherited ;
//PT3
if (param = 'S') then
   TFMul(Ecran).BOuvrir.Enabled := True;
//FIN PT3
//PT4
if ((param = 'D') and ((TFMul(Ecran).FListe.nbSelected <> 0) or
   (TFMul(Ecran).FListe.AllSelected=True))) then
   TFMul(Ecran).BOuvrir.Enabled := True;
//FIN PT4
//PT8
if ((param = 'U') and (Delete <> NIL) and
   ((TFMul(Ecran).FListe.nbSelected <> 0) or
   (TFMul(Ecran).FListe.AllSelected=True))) then
   Delete.Enabled := True;
//FIN PT8

MajTypeD;         //PT14
end;

procedure TOF_PGMULDADSHONOR.OnLoad;
var
QAnnee:TQuery;
begin
Inherited ;
DebExer:=IDate1900;
FinExer:=IDate1900;
QAnnee:= OpenSQL ('SELECT PEX_DATEDEBUT, PEX_DATEFIN'+
                  ' FROM EXERSOCIAL WHERE'+
                  ' PEX_EXERCICE="'+THAnnee.Value+'"',True);
if not QAnnee.eof then
   begin
   DebExer:=QAnnee.FindField('PEX_DATEDEBUT').AsDateTime;
   FinExer:=QAnnee.FindField('PEX_DATEFIN').AsDateTime;
   end;
Ferme(QAnnee);
DebExer2:=IDate1900;
FinExer2:=IDate1900;
QAnnee:= OpenSQL ('SELECT PEX_DATEDEBUT, PEX_DATEFIN, PEX_ANNEEREFER'+
                  ' FROM EXERSOCIAL WHERE'+
                  ' PEX_EXERCICE="'+ComboAnnee2.Value+'"',True);
if not QAnnee.eof then
   begin
   DebExer2:= QAnnee.FindField ('PEX_DATEDEBUT').AsDateTime;
   FinExer2:= QAnnee.FindField ('PEX_DATEFIN').AsDateTime;
   PGExercice2:= QAnnee.FindField ('PEX_ANNEEREFER').AsString; //PT16
   end;
Ferme(QAnnee);
ActiveWhere(Nil);
end;

procedure TOF_PGMULDADSHONOR.OnArgument(S: string);
var
AnneeE, AnneeCours, AnneePrec, arg, Buffer, ComboExer1, ComboExer2 : string;
Donnee, DonneeAffich, MoisE, StHonor, StPlus : string;
JourJ : TDateTime;
AnneeA, Jour, MoisM : Word;
TT : TFMul;
LibAnnee1, LibAnnee2 : THLabel;
TrouvePrec, TrouveCours : boolean;
QHonor : TQuery;
THonor, THonorD : TOB;
DbleCote, i, Long : integer;
begin
inherited;
arg:= S;
param:= Trim (ReadTokenPipe (arg, ';'));
TT:= TFMul (Ecran);
{$IFNDEF EAGLCLIENT}
Liste:= THDBGrid (GetControl ('FListe'));
{$ELSE}
Liste:= THGrid (GetControl ('FListe'));
{$ENDIF}
LibAnnee1:= THLabel (GetControl ('LANNEE'));
LibAnnee2:= THLabel (GetControl ('LANNEE1'));
ComboAnnee2:= THValComboBox (GetControl ('ANNEE1'));

TFMul(Ecran).OnKeyDown:= FormKeyDown;

if (param = 'S') then
   begin
   TFMul (Ecran).Caption:= 'Saisie - Consultation des honoraires';
   TFMul (Ecran).BOuvrir.OnClick:= GrilleDblClick;
   TFMul (Ecran).BInsert.OnClick:= NewHonor;
   if Liste <> nil then
      begin
{$IFNDEF EAGLCLIENT}
      Liste.MultiSelection:= False;
{$ENDIF}
      Liste.OnDblClick:= GrilleDblClick;
      end;
   end
else
if (param = 'D') then
   begin
   TFMul (Ecran).Caption:= 'Duplication des honoraires';
   TFMul (Ecran).BOuvrir.OnClick:= Duplication;
   TFMul (Ecran).BOuvrir.Enabled:= False;
   TFMul (Ecran).BInsert.Visible:= False;
{$IFNDEF EAGLCLIENT}
   TFMul (Ecran).FListe.MultiSelection:= True;
{$ENDIF}
   TFMul (Ecran).bSelectAll.Visible:= True;
   if (LibAnnee1 <> nil) then
      LibAnnee1.Caption:= 'De l''année';
   if (LibAnnee2 <> nil) then
      LibAnnee2.Visible:= True;
   if (ComboAnnee2 <> nil) then
      begin
      ComboAnnee2.Visible:= True;
      ComboAnnee2.OnChange:= Annee2Change;
      end;
   if (Liste <> nil) then
      begin
      Liste.OnFlipSelection:= GrilleClick;
      TFMul (Ecran).bSelectAll.OnClick:= GrilleClick;
      end;
   SetControlVisible ('CHCOMPL1', True); //PT14
   end
else
if (param = 'U') then
   begin
   TFMul (Ecran).Caption:= 'Suppression des honoraires';
   Delete:= TToolbarButton97 (GetControl ('BDELETE'));
   if (Delete <> nil) then
      begin
      Delete.Visible:= True;
      Delete.Enabled:= False;
      Delete.OnClick:= DeleteClick;
      end;
   TFMul (Ecran).bSelectAll.Visible:= True;
   TFMul (Ecran).BOuvrir.Visible:= False;
   TFMul (Ecran).BInsert.Visible:= False;
{$IFNDEF EAGLCLIENT}
   TFMul (Ecran).FListe.MultiSelection:= True;
{$ENDIF}
   if (Liste <> nil) then
      begin
      Liste.OnFlipSelection:= GrilleClick;
      TFMul (Ecran).bSelectAll.OnClick:= GrilleClick;
      end;
   end;

if (TT <> nil) then
   UpdateCaption (TT);
THAnnee:= THValCombobox (GetControl ('ANNEE'));
JourJ:= Date;
DecodeDate (JourJ, AnneeA, MoisM, Jour);
if (param = 'D') then
   begin
   if (MoisM > 9) then
      begin
      AnneePrec:= IntToStr(AnneeA - 1);
      AnneeCours:= IntToStr(AnneeA);
      end
   else
      begin
      AnneePrec:= IntToStr(AnneeA - 2);
      AnneeCours:= IntToStr(AnneeA - 1);
      end;
   end
else
   begin
   if (MoisM > 9) then
      begin
      AnneePrec:= IntToStr(AnneeA);
      AnneeCours:= IntToStr(AnneeA + 1);
      end
   else
      begin
      AnneePrec:= IntToStr(AnneeA - 1);
      AnneeCours:= IntToStr(AnneeA);
      end;
   end;

TrouvePrec:= RendExerSocialPrec (MoisE, AnneeE, ComboExer1, DebExer, FinExer,
                                 AnneePrec);
if (TrouvePrec = FALSE) then
   PGIBox ('L''exercice '+AnneePrec+' n''existe pas', Ecran.Caption);

TrouveCours:= RendExerSocialPrec (MoisE, AnneeE, ComboExer2, DebExer2, FinExer2,
                                  AnneeCours);
if ((TrouveCours = FALSE) and (param = 'D')) then
   PGIBox ('L''exercice '+AnneeCours+' n''existe pas', Ecran.Caption);

if THAnnee <> nil then
   begin
   if (TrouvePrec = TRUE) then
      THAnnee.Value:= ComboExer1
   else
      begin
      if (TrouveCours = TRUE) then
         THAnnee.Value:= ComboExer2;
      end;
   end;

if (ComboAnnee2 <> nil) then
   begin
   if (TrouveCours = TRUE) then
      ComboAnnee2.Value:= ComboExer2
   else
      begin
      if (TrouvePrec = TRUE) then
         ComboAnnee2.Value:= ComboExer1;
      end;
   end;

if (THAnnee <> nil) then
   begin
   PGAnnee:= THAnnee.value;
   PGExercice:= RechDom ('PGANNEESOCIALE', PGAnnee, False);
   PGExercice2:= RechDom ('PGANNEESOCIALE', ComboExer2, False);       //PT16
   end;

//PT16
if (VH_Paie.PGDecalage) then
   Decalage:= '03'
else
   Decalage:= '01';
//FIN PT16

Q_Mul:= TFMul (Ecran).Q;
TFMul (Ecran).SetDBListe ('PGSALDADS');

WW:= THEdit (GetControl ('XX_WHERE'));
BCherche:= TToolbarButton97 (GetControl ('BCherche'));

//PT12-1
StHonor:= 'SELECT * FROM DADSDETAIL WHERE PDS_SALARIE="--HONORAIRES--"';
if (ExisteSQL (StHonor)) then
   begin
   QHonor:= OpenSQL (StHonor, True);
   THonor:= TOB.Create ('DADSDETAIL', NIL, -1);
   THonor.LoadDetailDB('DADSDETAIL', '', '', QHonor, False);
   Ferme(QHonor);

   For i:=0 to (THonor.FillesCount (1)-1) do
       begin
       THonorD:= THonor.Detail.Items [i];
//PT13
       Donnee:= THonorD.GetValue ('PDS_DONNEE');
       DonneeAffich:= THonorD.GetValue ('PDS_DONNEEAFFICH');
       DbleCote:= Pos ('"', Donnee);
{PT15
       Long:= Length (Donnee);
       if (DbleCote <> 0) then
          Donnee:= Copy (Donnee, 1, DbleCote)+'"'+
                   Copy (Donnee, DbleCote+1, Long);
}
       Buffer:= Donnee;
       Donnee:= '';
       Long:= Length (Buffer);
       while (DbleCote <> 0) do
             begin
             Donnee:= Donnee+Copy (Buffer, 1, DbleCote)+'"';
             Buffer:= Copy (Buffer, DbleCote+1, Long);
             Long:= Length (Buffer);
             DbleCote:= Pos ('"', Buffer);
             end;
       Donnee:= Donnee+Copy (Buffer, DbleCote+1, Long);
//FIN PT15
       DbleCote:= Pos ('"', DonneeAffich);
{PT15
       Long:= Length (DonneeAffich);
       if (DbleCote <> 0) then
          DonneeAffich:= Copy (DonneeAffich, 1, DbleCote)+'"'+
                         Copy (DonneeAffich, DbleCote+1, Long);
}
       Buffer:= DonneeAffich;
       DonneeAffich:= '';
       Long:= Length (Buffer);
       while (DbleCote <> 0) do
             begin
             DonneeAffich:= DonneeAffich+Copy (Buffer, 1, DbleCote)+'"';
             Buffer:= Copy (Buffer, DbleCote+1, Long);
             Long:= Length (Buffer);
             DbleCote:= Pos ('"', Buffer);
             end;
       DonneeAffich:= DonneeAffich+Copy (Buffer, DbleCote+1, Long);
//FIN PT15
       StHonor:= 'INSERT INTO DADSDETAIL (PDS_SALARIE, PDS_TYPE, PDS_ORDRE,'+
                 ' PDS_DATEDEBUT, PDS_DATEFIN, PDS_EXERCICEDADS, PDS_ORDRESEG,'+
                 ' PDS_SEGMENT, PDS_DONNEE, PDS_DONNEEAFFICH)'+
                 ' VALUES'+
                 ' ("--H'+IntToStr (THonorD.GetValue ('PDS_ORDRE'))+'",'+
                 ' "'+THonorD.GetValue ('PDS_TYPE')+'",'+
                 ' '+IntToStr (THonorD.GetValue ('PDS_ORDRE'))+','+
                 ' "'+UsDateTime (THonorD.GetValue ('PDS_DATEDEBUT'))+'",'+
                 ' "'+UsDateTime (THonorD.GetValue ('PDS_DATEFIN'))+'",'+
                 ' "'+THonorD.GetValue ('PDS_EXERCICEDADS')+'",'+
                 ' '+IntToStr (THonorD.GetValue ('PDS_ORDRESEG'))+','+
                 ' "'+THonorD.GetValue ('PDS_SEGMENT')+'",'+' "'+Donnee+'",'+
                 ' "'+DonneeAffich+'")';
//FIN PT13
       ExecuteSQL(StHonor) ;
       end;

   FreeAndNil (THonor);
   StHonor:= 'DELETE FROM DADSDETAIL WHERE PDS_SALARIE="--HONORAIRES--"';
   ExecuteSQL (StHonor);
   end;

StHonor:= 'SELECT * FROM DADSPERIODES WHERE PDE_SALARIE="--HONORAIRES--"';
if (ExisteSQL (StHonor)) then
   begin
   QHonor:= OpenSQL (StHonor, True);
   THonor:= TOB.Create ('DADSPERIODES', NIL, -1);
   THonor.LoadDetailDB('DADSPERIODES', '', '', QHonor, False);
   Ferme(QHonor);

   For i:=0 to (THonor.FillesCount (1)-1) do
       begin
       THonorD:= THonor.Detail.Items [i];
       StHonor:= 'INSERT INTO DADSPERIODES (PDE_SALARIE, PDE_TYPE, PDE_ORDRE,'+
                 ' PDE_DATEDEBUT, PDE_DATEFIN, PDE_EXERCICEDADS, PDE_MOTIFDEB,'+
                 ' PDE_MOTIFFIN, PDE_DECALEE)'+
                 ' VALUES'+
                 ' ("--H'+IntToStr (THonorD.GetValue ('PDE_ORDRE'))+'",'+
                 ' "'+THonorD.GetValue ('PDE_TYPE')+'",'+
                 ' '+IntToStr (THonorD.GetValue ('PDE_ORDRE'))+','+
                 ' "'+UsDateTime (THonorD.GetValue ('PDE_DATEDEBUT'))+'",'+
                 ' "'+UsDateTime (THonorD.GetValue ('PDE_DATEFIN'))+'",'+
                 ' "'+THonorD.GetValue ('PDE_EXERCICEDADS')+'",'+
                 ' "'+THonorD.GetValue ('PDE_MOTIFDEB')+'",'+
                 ' "'+THonorD.GetValue ('PDE_MOTIFFIN')+'",'+
                 ' "'+THonorD.GetValue ('PDE_DECALEE')+'")';
       ExecuteSQL(StHonor) ;
       end;

   FreeAndNil (THonor);
   StHonor:= 'DELETE FROM DADSPERIODES WHERE PDE_SALARIE="--HONORAIRES--"';
   ExecuteSQL (StHonor);
   end;
//FIN PT12-1

ActiveWhere (nil);
//PT14
Compl:= TCheckBox (GetControl ('CHCOMPL'));
if (Compl <> nil) then
   Compl.OnClick:= ComplClick;
//FIN PT14

//PT17
StPlus:= ' WHERE ET_FICTIF<>"X"';
SetControlProperty ('ETAB', 'Plus', StPlus);
//FIN PT17
end;

procedure TOF_PGMULDADSHONOR.OnClose;
begin
inherited;
end;

//PT3
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/07/2002
Modifié le ... :   /  /
Description .. : Double-click sur la grille
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSHONOR.GrilleDblClick(Sender: TObject);
var
Ordre, StSal, Sal, TypeDe : String;
Existe : boolean;
begin
//PT7
if ((Q_Mul <> nil) and (Q_Mul.RecordCount = 0)) then
   exit;
//FIN PT7
Sal:= '--H';

//PT14
Ordre:= IntToStr (Q_Mul.FindField ('PDE_ORDRE').AsInteger);
if (Copy (TypeD, 1, 1)='2') then
   begin
   StSal:= 'SELECT PDS_SALARIE'+
           ' FROM DADSDETAIL WHERE'+
           ' PDS_ORDRE='+Ordre+' AND'+
           ' PDS_SALARIE LIKE "'+Sal+'%" AND'+
           ' PDS_TYPE="0'+Copy (TypeD, 2, 2)+'" AND'+
           ' PDS_DATEDEBUT="'+UsDateTime (DebExer)+'" AND'+
           ' PDS_DATEFIN="'+UsDateTime (FinExer)+'" AND'+
           ' PDS_EXERCICEDADS = "'+PGExercice+'"';
   Existe:= ExisteSQL (StSal);
   end
else
   begin
   StSal:= 'SELECT PDS_SALARIE'+
           ' FROM DADSDETAIL WHERE'+
           ' PDS_ORDRE='+Ordre+' AND'+
           ' PDS_SALARIE LIKE "'+Sal+'%" AND'+
           ' PDS_TYPE="2'+Copy (TypeD, 2, 2)+'" AND'+
           ' PDS_DATEDEBUT="'+UsDateTime (DebExer)+'" AND'+
           ' PDS_DATEFIN="'+UsDateTime (FinExer)+'" AND'+
           ' PDS_EXERCICEDADS = "'+PGExercice+'"';
   Existe:= ExisteSQL (StSal);
   end;

if (Existe=False) then
   begin
//FIN PT14
   if (param = 'S') then
      begin
{PT14
      Ordre:= IntToStr(Q_Mul.FindField ('PDE_ORDRE').AsInteger);
}
      StSal:= 'SELECT PDE_SALARIE'+
              ' FROM DADSPERIODES WHERE'+
              ' PDE_ORDRE='+Ordre+' AND'+
              ' PDE_SALARIE LIKE "'+Sal+'%"';
{$IFNDEF EAGLCLIENT}
      TheMulQ:= THQuery(Ecran.FindComponent('Q'));
{$ELSE}
      TheMulQ:= TOB(Ecran.FindComponent('Q'));
{$ENDIF}
      if PGAnnee <> '' then
         begin
         if ((ExisteSQL(StSal)=FALSE) and (PGAnnee <> '')) then
            AGLLanceFiche ('PAY','DADS_HONORAIRES',  '',Ordre+';'+Sal+Ordre+';1',
                           'CREATION;'+Ordre+';'+PGAnnee)
         else
            AGLLanceFiche ('PAY','DADS_HONORAIRES',  '',Ordre+';'+Sal+Ordre+';1',
                           'MODIFICATION;'+Ordre+';'+PGAnnee);
         end
      else
         PGIBox ('L''année n''est pas valide','DADS-U');
      end;
//PT14
   end
else
   begin
   if (Copy (TypeD, 1, 1)='2') then
      TypeDe:= 'normale'
   else
      TypeDe:= 'complémentaire';
   PGIBox ('Traitement impossible : L'' honoraire '+Ordre+' a été#13#10'+
           'intégré dans une déclaration '+TypeDe, TFMul (Ecran).Caption);
   end;
//FIN PT14

//PT10
if BCherche<>nil then
   BCherche.click;
//FIN PT10
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/07/2002
Modifié le ... :   /  /
Description .. : Click sur la grille
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSHONOR.GrilleClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
Liste:= THDBGrid(GetControl('FListe'));
{$ELSE}
Liste:= THGrid(GetControl('FListe'));
{$ENDIF}
if Sender = TFMul(Ecran).bSelectAll then
   Liste.AllSelected:= not Liste.AllSelected ;

if (param = 'D') then
   begin
   if Liste <> NIL then
      BEGIN
      if (Liste.NbSelected<>0) or (Liste.AllSelected) then
         TFMul(Ecran).BOuvrir.Enabled:= True
      else
         TFMul(Ecran).BOuvrir.Enabled:= False;
      END;
   end;

//PT8
if (param = 'U') then
   begin
   if ((Liste <> NIL) and (Delete <> NIL)) then
      BEGIN
      if (Liste.NbSelected<>0) or (Liste.AllSelected) then
         Delete.Enabled:= True
      else
         Delete.Enabled:= False;
      END;
   end;
//FIN PT8
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/07/2002
Modifié le ... :
Description .. : XX_WHERE
Mots clefs ... :
*****************************************************************}
procedure TOF_PGMULDADSHONOR.ActiveWhere(Sender: TObject);
var
Etab, StEtab, Where : string;
TEtabDADS, TEtabDADSD : tob;
begin
PGAnnee:= THAnnee.value;
PGExercice:= RechDom ('PGANNEESOCIALE', PGAnnee, False);
if (WW <> NIL) AND (THAnnee <> NIL) then
   begin
   Where:= ' PDE_SALARIE LIKE "--H%" AND'+
           ' (PDS_SEGMENT="S70.G01.00.001" OR'+
           ' PDS_SEGMENT="S70.G01.00.002.001") AND'+
           ' PDE_ORDRE=PDS_ORDRE AND'+
           ' PDE_DATEDEBUT="'+UsDateTime(DebExer)+'" AND'+
           ' PDS_DATEDEBUT="'+UsDateTime(DebExer)+'" AND'+
           ' PDE_TYPE="'+TypeD+'"';

   Etab:= GetControlText('ETAB');
   If GetControlText('ETAB') <>'' Then
      SetControlText ('XX_WHERE', Where+ ' AND PDS_ORDRE IN'+
                      ' (SELECT DISTINCT PDS_ORDRE FROM DADSDETAIL WHERE'+
                      ' PDS_SEGMENT="S70.G01.00.014" AND'+
                      ' PDS_DONNEEAFFICH="'+Etab+'")')
   Else
{PT17
      SetControlText ('XX_WHERE', where);
}
      begin
      StEtab:= 'AND (';
      TEtabDADS:= TOB.Create ('Les etablissements', nil, -1);
      ChargeEtabNonFictif (TEtabDADS);
      if (TEtabDADS<>nil) then
         begin
         TEtabDADSD:= TEtabDADS.FindFirst ([''], [''], False);
         if (TEtabDADSD<>nil) then
            Etab:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT')
         else
            StEtab:= '';
         while (Etab<>'') do
               begin
               StEtab:= StEtab+' PDS_DONNEEAFFICH="'+Etab+'"';
               TEtabDADSD:= TEtabDADS.FindNext ([''], [''], False);
               if (TEtabDADSD<>nil) then
                  Etab:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT')
               else
                  Etab:='';
               if (Etab<>'') then
                  StEtab:= StEtab+' OR'
               else
                  StEtab:= StEtab+')';
               end;
         end
      else
         StEtab:= '';
      FreeAndNil (TEtabDADS);
      if (StEtab<>'') then
         SetControlText ('XX_WHERE', Where+ ' AND PDS_ORDRE IN'+
                         ' (SELECT DISTINCT PDS_ORDRE FROM DADSDETAIL WHERE'+
                         ' PDS_SEGMENT="S70.G01.00.014" '+StEtab+')')
      else
         SetControlText ('XX_WHERE', Where);
      end;
//FIN PT17

   TFMul(Ecran).BOuvrir.Enabled:= False;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/07/2002
Modifié le ... :   /  /
Description .. : Duplication des honoraires d'une année sur l'autre
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSHONOR.Duplication(Sender: TObject);
var
i, OrdreDest : integer;
StHonor : string;
QHonor:TQuery;
Maintenant : TDateTime;
begin
StHonor := '--H';
{$IFNDEF EAGLCLIENT}
Liste:= THDBGrid(GetControl('FListe'));
{$ELSE}
Liste:= THGrid(GetControl('FListe'));
{$ENDIF}
if Liste <> NIL then
   BEGIN
   if (Liste.NbSelected=0) and (not Liste.AllSelected) then
      begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
      end;

   Trace:= TStringList.Create;                        //PT5
   Maintenant:= Now;
   if Trace <> nil then
      Trace.Add ('Début de duplication à '+TimeToStr (Maintenant));

   OrdreDest:= 1;
   QHonor:= OpenSQL('SELECT MAX(PDE_ORDRE) AS NUMMAX FROM DADSPERIODES WHERE'+
                    ' PDE_SALARIE LIKE "'+StHonor+'%"',True);
   If Not QHonor.eof then
      try
      OrdreDest:= (QHonor.FindField('NUMMAX').AsInteger)+1;
      except
            on E: EConvertError do
               OrdreDest:= 1;
      end;
   Ferme(QHonor);

   if (Liste.AllSelected=TRUE) then
      begin
      InitMoveProgressForm (NIL, 'Duplication en cours',
                            'Veuillez patienter SVP ...',
                            TFmul(Ecran).Q.RecordCount,FALSE,TRUE);
      InitMove (TFmul(Ecran).Q.RecordCount,'');

{$IFDEF EAGLCLIENT}
      if (TFMul(Ecran).bSelectAll.Down) then
         TFMul(Ecran).Fetchlestous;
{$ENDIF}

      TFmul(Ecran).Q.First;
      while Not TFmul(Ecran).Q.EOF do
            BEGIN
            Duplique_un(OrdreDest);
            TFmul(Ecran).Q.Next;
            END;

      Liste.AllSelected:= False;
      TFMul(Ecran).bSelectAll.Down:= Liste.AllSelected;
      end
   else
      begin
      InitMoveProgressForm (NIL, 'Duplication en cours',
                            'Veuillez patienter SVP ...', Liste.NbSelected,
                            FALSE, TRUE);
      InitMove(Liste.NbSelected,'');

      for i:=0 to Liste.NbSelected-1 do
          BEGIN
          Liste.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
          TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
{$ENDIF}
          Duplique_un(OrdreDest);
          END;

      Liste.ClearSelected;
      end;

   FiniMove;
   FiniMoveProgressForm;
   TFMul(Ecran).BOuvrir.Enabled:= False;
   PGIBox ('Traitement terminé', 'Duplication des honoraires');
   Maintenant:= Now;
   if Trace <> nil then
      Trace.Add ('Duplication terminée à '+TimeToStr (Maintenant));

//PT5
{$IFNDEF DADSUSEULE}
   CreeJnalEvt ('001', '045', 'OK', NIL, NIL, Trace);
{$ENDIF}
   FreeAndNil (Trace);
//FIN PT5
   END;
if BCherche<>nil then
   BCherche.click;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/05/2004
Modifié le ... :   /  /
Description .. : Duplication d'un seul honoraire
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSHONOR.Duplique_un(var OrdreDest: integer);
var
OrdreHonor : integer;
Maintenant : TDateTime;
LeType : string;
begin
OrdreHonor:= TFmul(Ecran).Q.FindField('PDE_ORDRE').asinteger;
try
   begintrans;
   ChargeTOBHon (OrdreHonor);
//PT14
   LeType:= TypeD;
   if (GetControlText ('CHCOMPL1')='X') then
      TypeD:= '201'
   else
      TypeD:= '001';
//FIN PT14
   DeleteErreur ('', 'SKO');	//PT12-2
   Duplic (OrdreDest);
   EcrireErreurKO;	//PT12-2
   LibereTOBHon;
   TypeD:= LeType;      //PT14
   CommitTrans;
Except
   Rollback;
   Maintenant:= Now;
   if Trace <> nil then
      Trace.Add ('Honoraire '+IntToStr (OrdreHonor)+' :'+
                 ' Duplication annulée à '+TimeToStr (Maintenant));
   end;

MoveCur (False);
MoveCurProgressForm (IntToStr(OrdreHonor));
OrdreDest:= OrdreDest+1;
end;

//PT8
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 26/02/2004
Modifié le ... :   /  /
Description .. : Suppression des honoraires
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSHONOR.DeleteClick(Sender: TObject);
var
i, reponse : integer;
Maintenant : TDateTime;
begin
{$IFNDEF EAGLCLIENT}
Liste:= THDBGrid(GetControl('FListe'));
{$ELSE}
Liste:= THGrid(GetControl('FListe'));
{$ENDIF}

reponse:= PGIAsk('Cette commande supprimera les honoraires sélectionnées.#13#10'+
                 'Voulez-vous continuer ?', TFMul(Ecran).Caption);
if (reponse <> mrYes) then
   exit;

if Liste <> NIL then
   BEGIN
   if (Liste.NbSelected=0) and (not Liste.AllSelected) then
      begin
      MessageAlerte ('Aucun élément sélectionné');
      exit;
      end;

   Trace:= TStringList.Create;
   if (Liste.AllSelected=TRUE) then
      begin
      InitMoveProgressForm (NIL, 'Suppression en cours',
                            'Veuillez patienter SVP ...',
                            TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
      InitMove (TFmul(Ecran).Q.RecordCount,'');

      Maintenant:= Now;
      if Trace <> nil then
         Trace.Add ('Début de suppression à '+TimeToStr (Maintenant));
{$IFDEF EAGLCLIENT}
      if (TFMul(Ecran).bSelectAll.Down) then
         TFMul(Ecran).Fetchlestous;
{$ENDIF}
      TFmul(Ecran).Q.First;
      while Not TFmul(Ecran).Q.EOF do
            BEGIN
            Delete_un;
            TFmul(Ecran).Q.Next;
            END;

      Liste.AllSelected:= False;
      TFMul(Ecran).bSelectAll.Down:= Liste.AllSelected;
      end
   else
      begin
      InitMoveProgressForm (NIL, 'Calcul en cours',
                            'Veuillez patienter SVP ...',
                            Liste.NbSelected, FALSE, TRUE);

      InitMove(Liste.NbSelected,'');
      Maintenant:= Now;
      if Trace <> nil then
         Trace.Add ('Début de suppression à '+TimeToStr (Maintenant));

      for i:=0 to Liste.NbSelected-1 do
          BEGIN
          Liste.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
          TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
{$ENDIF}
          Delete_un;
          END;

      Liste.ClearSelected;
      end;

   FiniMove;
   FiniMoveProgressForm;
   if (Delete <> NIL) then
      Delete.Enabled:= False;

   PGIBox ('Traitement terminé', TFMul(Ecran).Caption);
   Maintenant:= Now;
   if Trace <> nil then
      Trace.Add ('Suppression terminée à '+TimeToStr (Maintenant));

{$IFNDEF DADSUSEULE}
   CreeJnalEvt ('001', '046', 'OK', NIL, NIL, Trace);
{$ENDIF}
   FreeAndNil (Trace);
   END;
if BCherche<>nil then
   BCherche.click;
end;
//FIN PT8

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/05/2004
Modifié le ... :   /  /
Description .. : Suppression d'un seul honoraire
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSHONOR.Delete_un();
var
OrdreHonor : integer;
StHonor : string;
Maintenant : TDateTime;
begin
StHonor:= '--H';
OrdreHonor:= TFmul(Ecran).Q.FindField('PDE_ORDRE').asinteger;
try
   begintrans;
   if Trace <> nil then
      Trace.Add ('Traitement de l''honoraire '+IntToStr (OrdreHonor));
{PT14
   DeletePeriode (StHonor+IntToStr (OrdreHonor), '001', OrdreHonor);
   DeleteDetail (StHonor+IntToStr (OrdreHonor), '001', OrdreHonor);
}
   DeletePeriode (StHonor+IntToStr (OrdreHonor), OrdreHonor);
   DeleteDetail (StHonor+IntToStr (OrdreHonor), OrdreHonor);
//FIN PT14   
   CommitTrans;
Except
   Rollback;
   Maintenant:= Now;
   if Trace <> nil then
      Trace.Add ('Honoraire '+IntToStr (OrdreHonor)+' :'+
                 ' Suppression annulée à '+TimeToStr (Maintenant));
   end;
MoveCur (False);
MoveCurProgressForm (IntToStr (OrdreHonor));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 15/07/2002
Modifié le ... :   /  /
Description .. : Création d'un honoraire
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSHONOR.NewHonor(Sender: TObject);
var
StHonor : String;
OrdreDest : integer;
QHonor : TQuery;
begin
OrdreDest:= 1;
StHonor:= '--H';
QHonor:= OpenSQL ('SELECT MAX(PDE_ORDRE) AS NUMMAX FROM DADSPERIODES WHERE'+
                  ' PDE_SALARIE LIKE "'+StHonor+'%"',True);
if not QHonor.eof then
   try
   OrdreDest:= (QHonor.FindField ('NUMMAX').AsInteger)+1;
   except
         on E: EConvertError do
            OrdreDest:= 1;
   end;
Ferme(QHonor);
AGLLanceFiche ('PAY','DADS_HONORAIRES','',
               IntToStr(OrdreDest)+';'+StHonor+IntToStr(OrdreDest)+';1',
               'CREATION;'+IntToStr(OrdreDest)+';'+PGAnnee);
if BCherche<>nil then
   BCherche.click;
end;
//FIN PT3

//PT4
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/10/2002
Modifié le ... :   /  /
Description .. : Désactivation de bOuvrir dans le cas ou l'annee 2 change
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSHONOR.Annee2Change(Sender: TObject);
begin
TFMul(Ecran).BOuvrir.Enabled:=False;
end;
//FIN PT4

//PT9
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 04/05/2004
Modifié le ... :   /  /
Description .. : Complément des raccourcis claviers
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSHONOR.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
if (param = 'D') then
   case Key of
        VK_F6: if ((TFMul(Ecran).BOuvrir.Visible) and
                  (TFMul(Ecran).BOuvrir.Enabled)) then
                  TFMul(Ecran).BOuvrir.Click; //Duplication des éléments
        end;

if (param = 'U') then
   case Key of
        VK_Delete: if ((GetControlVisible('BDELETE')) and
                      (GetControlEnabled('BDELETE'))) then
                      Delete.Click; //Suppression des éléments
        end;
end;
//FIN PT9

//PT14
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/11/2006
Modifié le ... :   /  /
Description .. : Clic sur case à cocher "Déclaration complémentaire"
Mots clefs ... :
*****************************************************************}
procedure TOF_PGMULDADSHONOR.ComplClick (Sender: TObject);
begin
MajTypeD;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/11/2006
Modifié le ... :   /  /
Description .. : Mise à jour du type de déclaration
Mots clefs ... :
*****************************************************************}
procedure TOF_PGMULDADSHONOR.MajTypeD ();
begin
if (GetControlText ('CHCOMPL')='X') then
   TypeD:= '201'
else
   TypeD:= '001';
end;
//FIN PT14

initialization
  registerclasses([TOF_PGMULDADSHONOR]);
end.

