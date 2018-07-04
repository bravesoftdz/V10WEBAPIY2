unit UTILEDT;

//================================================================================
// Interface
//================================================================================
interface

uses
    SysUtils,
    WinTypes,
    WinProcs,
    Messages,
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    StdCtrls,
    Spin,
{$IFDEF EAGLCLIENT}
    UTOB,
{$ELSE}
    DB,
   {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    HQuickRP,
    {$IFDEF EDTQR}
    QRRupt,
    {$ENDIF}
  {$IFDEF V530}
    EdtDoc,
  {$ELSE}
    EdtRDoc,
  {$ENDIF}
    EdtQR,
{$ENDIF}
    Hctrls,
    FileCtrl,
    ExtCtrls,
    HStatus,
    ComCtrls,
    Ent1,
    HEnt1,
    Buttons,
    HCompte,
    Printers,
    CritEdt,
    Mask,
    {$IFDEF MODENT1}
    CPTypeCons,
    {$ELSE}
    tCalcCum,
    {$ENDIF MODENT1}
    ULibExercice  ,UentCommun
    ;

{$IFDEF EAGLCLIENT}
    Type TQuery = TOB;
{$ENDIF}

{ Méthode à suivre :
1 - Enrichir le TCRITXXX avec
  Formatprint : TFormatPrint pour les balances
  FormatprintEDT : TFormatPrintEDT pour les autres.

Attention  : Pour les grands livres, utiliser le type TCRITGL au lieu du type TCRITBAL

2 - Mettre à jour le multi critère au niveau visuel sur l'onglet "option d'édition"
(Colonnage et ligne, prévisualisation, 2 en 1, ...)

  Attention : Normes à suivre :
  Le checkbox de la "liste de prévisualisation" s'appelle FLISTE
  Le checkbox de la "Deux page ur une" s'appelle REDUIRE

  Attention : Ordre de tabulation entre champ et panels

3 - Colonner les composants via les "Tags" (Un Colonne = 1 Champs)
    Indiquer la propriété SYNTITRE à TRUE sur les titres
    Sur le 1° Composant Avec SYNTITRE=TRUE (Et donc Tag=1) Mettre TAG=-1


4 - Implémenter la band OVERLAY avec les traits de séparation
==> QRLIGNE En bottomBad et topBand pour les traits
==> QRLIGNE en horizBand pour les lignes de séparation(TAG !!!)

5 - Implémenter la récupération du multicritère pour CRITXXX.formatPrint dans RECUPCRITXXX

  Attention : pour les éditions de type balance, rajouter :
      PutTabCol(X,TabColl[X],TitreColCpt.Tag,TRUE ou FALSE) ;
      ( permet de dire si la colonne X s'affiche ou non )


6 - Mettre à jour INITDIVERS :
==> Pour les Separateur de ligne (lien avec le multicritère)

7 - Déclarer un PremTabCol Dans la form de type TTabCollAff pour les éditions de type balances
et TTabCollAffEdt pour les autres éditions.

8 - Implémenter dans FORMSHOW :
==> InitFormatBal(Self,PremTabCol) ;
==> InitEdit(EntetePage,QRP) ; ( pour les éditions de type balance )

==> InitFormatEdt(Self,PremTabCol,) ;
==> InitEdit(EntetePage,QRP) ; ( pour les autres éditions )

9 - Implémenter dans INITDIVERS :
==>ll:=ChangeFormatBal(Self,PremTabCol,CritBal.FormatPrint.TabColl) ;
   CalculPourMiseEnPage(ll,QRP,Self,CritBal.FormatPrint) ; ( pour les éditions de type balance )
   et en plus, Pour les balances uniquement :
   ChangeFormatCompte(fbGene,Self,G_GENERAL.Width,1,2) ;

==>ll:=ChangeFormatBalEdt(Self,PremTabCol,CritBal.FormatPrint.TabColl) ;
   CalculPourMiseEnPageEdt(ll,QRP,Self,CritBal.FormatPrint) ;  ( pour les autres éditions )

10 - Implémenter dans BVALIDERCLICK apère appel à CHOIXEDITION :

SynOnGrid:=(FListe.Checked) ;
if Reduire.Checked then QRPrinter.Thumbs:=2 else QRPrinter.Thumbs:=1 ;


PASSAGE POUR APPEL HYPERZOOM :

1 : Rajouter dans la form 2 variables :

    OkMajEdt                         : Boolean ;
    OkZoomEdt                        : Boolean ;

2 : Dans La procedure principale :

 QR.OkMajEdt:=TRUE ;
 QR.OkZoomEdt:=FALSE ;

3 : Déclarer une procedure XXZOOM

Exemple : Pour GLGENERAL

procedure GLGeneralZoom(Crit : TCritGL) ;
var QR : TGdLivGen ;
begin
QR:=TGdLivGen.Create(Application) ;
try
 QRPrinter.OnSynZoom:=QR.GLGenZoom ;
 QR.OkMajEdt:=FALSE ;
 QR.OkZoomEdt:=TRUE ;
 QR.CritGL:=Crit ;
 QR.FListe.Checked:=FALSE ;
 QR.Reduire.Checked:=FALSE ;
 QR.Apercu.Checked:=TRUE ;
 InitFormatEdt(QR,QR.PremTabCol,QR.CritGL.FormatPrint.TabColl) ;
 InitEdit(QR.EntetePage,QR.QRP) ;
 QR.BValiderClick(Nil) ;
 finally
 QR.Free ;
 end;
end;

4 : Virer CtrlPerExo (Définition)

5 : Changer CRITXXOK

Exemple : Pour CRITGLOK

if Not OkZoomEdt then RecupCritGL ;
Result:=CtrlPerExo(CritGL.DateDeb,CritGL.DateFIn) ;
if Result then
   begin
   Dev:=CritGL.DeviseSelect<>'' ;
   Etab:=CritGL.Etab<>'' ;
   Exo:=CritGL.MonoExo And (CritGL.Exo.Code<>'') ;
   QBal:=PrepareTotCpt(fbGene,[],Dev,Etab,Exo) ;
   end;

6 : Mettre En début de Renseignecritere :

if OkZoomEdt then begin BTitre.Enabled:=FALSE ; Exit ; end else BTitre.Enabled:=FALSE ;

7 : Mettre En Fin de Renseignecritere :

InitQrPrinter(FListe.Checked,Reduire.Checked,FCouleur.Checked) ;

8 : Virer de BValiderClick :

   SynOnGrid[SynCurQRPrinter]:=(FListe.Checked) ;
   if Reduire.Checked then QRPrinter.Thumbs:=2 else QRPrinter.Thumbs:=1 ;

9 : Gerer un InitCritXX (voir InitCritGL En fin de UTILEDT)

10 : Mettre apres le preview dans BValiderClick :

   if OkMajEdt then
      MajEdition('XXX',CritBal.Exo.Code,DateToStr(CritBal.DateDeb),DateToStr(CritBal.DateFin),
                 '',TotCpt[3].TotDebit,TotCpt[3].TotCredit,TotCpt[4].TotDebit,TotCpt[4].TotCredit) ;

     Et dans le USes de l'implémentation une référence à EdtLegal.


}

const MaxColListe=10 ;

type  TCPTBCP = Class
      Deb, Fin  : Array[1..5] of String ;
      Total     : Array[0..12] of Double ;
      QuelCpt   : char ;
      end;

Type TabTot12 = Array[0..12] Of TabDC ;
Type TMontTot = Array[0..2] Of TabTot12 ;
Type TabMont77 = Array[0..77] of Double ;

//==================================================
// Definition des fonction
//==================================================
{ Pour les éditions de type balance : }
{$IFNDEF EAGLCLIENT}
procedure CALCULPOURMISEENPAGE(L : Integer ; QRP : TQuickReport ; FF : TForm ; Var CFP : TFormatPrint) ;
procedure InitFormatBal ( FF : TForm ; Var PremTabCol : TTabCollAff) ;
function  ChangeFormatBal ( FF : TForm ; Var PremTabCol,TabCol : TTabCollAff) : Integer ;
procedure ChangeFormatCompte ( fb1,fb2 : TFichierBase ; FF : TForm ; Lg1,Tag1,Tag2 : Integer ; NomCompte : String) ;
{$ENDIF}
procedure PutTabCol(i : Integer ; Var Coll : TCollAff ; Tag : Integer ; Affiche : Boolean) ;

{ Pour les autres éditions : }
{$IFNDEF EAGLCLIENT}
procedure CALCULPOURMISEENPAGEEDT(L : Integer ; QRP : TQuickReport ; FF : TForm ; Var CFP : TFormatPrintEdt) ;
procedure InitFormatEdt ( FF : TForm ; Var PremTabCol,TabCol : TTabCollAffEdt) ;
function  ChangeFormatEdt ( FF : TForm ; Var PremTabCol,TabCol : TTabCollAffEdt ; initTabCol : Boolean = TRUE) : Integer ;
{$ENDIF}

{ Commun à toutes les éditions : }
{$IFNDEF EAGLCLIENT}
procedure InitEdit(Entete : TQRBAND ; QRP : TQuickReport) ;
procedure InitQRPRINTER(QR : TQuickReport ; SurGrid,Reduire,Couleur : Boolean) ;
procedure ReInitEdit(Entete : TQRBAND ; AvecCouleur : Boolean) ;
procedure POSITIONNEFOURCHETTE(TC1,TC2 : TComponent ; QRL1,QRL2 : TQRLABEL) ;
function  GETPARAMLISTE(NewValue : String ; FF : TForm ) : String ;
procedure ALIMLISTE(Q : TQuery ; i : Integer ; FF : TForm) ;
{$ENDIF}
function  CtrlPerExo(Date1,Date2 : TDateTime) : Boolean ;
function  CtrlMouvementeSur(MouvementeSur : integer ; NatureEcr : String) : Boolean ;
procedure POSITIONNEFOURCHETTEST(TC1,TC2 : TComponent ; Var ST1,ST2 : String) ;
//procedure PREPARECREATETFIELD( Q : TQuery ; Var LLF,LLFD : TStringList) ;
//procedure CREATETFIELD(FF : TForm ; Q : TQuery ; SQLGA : String ; LLF,LLFD : TStringList) ;
procedure ADDREPORT(Seti : SetOfByte ; Var Report : TabReport ; D,C : Double ;Decimale : Integer ) ;
procedure INITREPORT(Seti : SetOfByte ; Var Report : TabReport) ;
function  QUELREPORT(Report : TabReport ; Var D,C : Double) : Byte;
procedure ADDREPORTBAL(Seti : SetOfByte ; Var Report : TabReport ; Tot : TabTRep ;Decimale : Integer ) ;
function  QUELREPORTBAL(Report : TabReport ; Var Tot : TabTRep) : Byte ;
procedure InitCritEdt (Var CritEdt : TcritEdt ) ;

{$IFNDEF EAGLCLIENT}
procedure CASEACOCHER(FBox : TCheckBox ; RBox : TQRLabel);
procedure DATEBANDEAUBALANCE(Avant,Selection,Apres,Solde : TQRLABEL ; CritEdt : TCritEdt ; St0,St1,St2 : String) ;
procedure DATECUMULAUGL(CumulAu : TQRLABEL ; CritEdt : TCritEdt ; St0 : String) ;
function AfficheNewMontant( Var CritEdt : TCritEdt ; TB : TabDC ; MontantLabel : TQRLabel) : Boolean ;
procedure ChgMaskQrCalc (C : TQRDBCalc ; Decim : Integer ; AffSymb : boolean ; Symbole : String ; IsSolde : Boolean) ;
{$ENDIF}
function  DEUX2UNCOMPTE(St1,St2 : String) : String ;

{ Pour les editions de grand livre }
procedure ChargeBilan(Var L : TStringList) ;
procedure AddBilan (L : TStringList ; LeCompte : String ; Montants : Array of double) ;
procedure PrintBilan (L : TStringList ; Var Sommes : Array of double ; Qui : Char) ;
procedure VideBilan(Var L : TStringList) ;
function  QuelleTypeRupt( i :Integer ; SAnsRupt,AvecRupt,SurRupt : Boolean ) : TRuptEtat ;
procedure DExoToDates ( Exo : String3 ; Var D1,D2 : TDateTime ) ;
procedure VoirQuelAN(ValueLTypCpt : String ; CodeExo : String ; FD : TMaskEdit ; FAvecAN : TCheckBox) ;
function  ValiQuali(Valide, Qualif : String) : String ;
function  ValideTva(Valide : String) : String ;

procedure IntervalleDateBalGL1(MonoExo : Boolean ; LExo : TExoDate ; Var D1,D11,D21 : TDateTime) ;
procedure IntervalleDateBalGL2(MonoExo : Boolean ; LExo : TExoDate ; Var D1,D11,D21 : TDateTime) ;

function  OrderLibre(LibreOrder : String ; SansVirgule : Boolean=FALSE) : String ;
function  WhereLibre(LesCodes1, LesCodes2 : String ; LeFb : TFichierBase ; OnlyCpteAssocie : Boolean) : string ;

{ Pour les éditions en Fiche et en Liste}
{$IFNDEF EAGLCLIENT}
procedure RedimLargeur(FF : TForm ; Qrp : TQuickReport ; Couleur : Boolean ; TTitre : TQRSysData);
procedure ReajusteLabels(BDCol : TQRBand);
{$ENDIF}

{ Pour éditions Budgétaires }
{$IFNDEF EAGLCLIENT}
procedure AffecteLigne(FF : TForm ; Ent : TQRBAND ;CFP : TFormatPrintEDT) ; {Creation de lignes verticaux}
procedure AfficheBudgetEn(RDevises : TQRLabel ; Var CritEdt : TCritEdt) ;
procedure FreeLesLignes(FF : TForm) ; {Suppression de ces mêmes lignes verticaux}
{$ENDIF}
function  InitDecimaleResol(Resol : hChar ; Decimale : Integer) : Integer ;
procedure SwapSymbolePivotDevise(Var CritEdt : TCritEdt) ;
procedure Reevaluation(Var D,C : Double ; Quelle : hString ; CoefRev : double ) ;
function  ExistBud(LeFb : TFichierBase ; Valeur, Jal, Axe : String ; First : Boolean ) : String ;
function  ExistBudOle(LeFb : TFichierBase ; Valeur, Jal, Axe : String ; First : Boolean ; AL1, AL2 : String ) : String ;
procedure InitResolution(FRESOLUTION: THValComboBox) ;
procedure StPourMontantsBudgetEcart(M : TabTot12 ; Var CritEdt : TCritEdt ; DebitPos : Boolean ; var AffM : array of String) ;
procedure AlimTotEdtEcartBudget(Var LeTotal : TabTot12 ; Var SousTot : TMontTot ; Var CritEdt : TCritEdt) ;

{$IFNDEF EAGLCLIENT}
{$IFDEF EDTQR}
procedure AlimRuptSup(HauteurBandeRuptIni,QuelleRupt,NewWidth : Integer ; BRupt : TQRBAND ; LibRuptInf : Array Of TRuptInf ; FF : TForm) ;
function  AlimStRuptSup(QuelleRupt : Integer ; Cod,Lib : String ; LibRuptInf : Array Of TRuptInf) : String ;
{$ENDIF}
{$ENDIF}

// récup. des racines / comptes coll. client pour TVA/Encaissement
function  SQLCollCliEncTVA ( Alias : String ) : String ;
function  SQLCollFouEncTVA ( Alias : String ) : String ;

{$IFNDEF EAGLCLIENT}
function  SautPageRuptAFaire(Var CritEdt : tCritEdt ; QRB : TQRBand ; QuelleRupt : Integer) : Boolean ;
{$ENDIF}
function  TrouveLibTP(St : String ; OkLib : Boolean) : String ;

{$IFNDEF EAGLCLIENT}
procedure LanceDoc(Tip,Nat,Modele : String ; TT : Tlist ; Load : TLoad_Etiq ; Apercu,PrintDialog : boolean) ;
{$ENDIF}

Procedure MajEditionLegal(LeType,Exo,D1,D2: String) ;

//================================================================================
// Implementation
//================================================================================
implementation

  {$IFDEF MODENT1}
uses
  CPProcMetier,
  CPProcGen;
  {$ENDIF MODENT1}


Type TCritAff = RECORD
                ColCrit : Integer ;
                DecalTop,DecalLeft : Integer ;
                end;

Type TabCritAff = Array[1..3] Of TCritAff ;

(*======================================================================*)
procedure IntervalleDateBalGL1(MonoExo : Boolean ; LExo : TExoDate ; Var D1,D11,D21 : TDateTime) ;
Var DD1 : TDateTime ;
    Exo : TExoDate ;
begin
DD1:=D1 ;
if MonoExo then Exo:=LExo else Exo:=VH^.EnCours ;
if Exo.Deb=D1 then begin D11:=D1 ; D21:=D11 ; end else begin D1:=Exo.Deb ; D11:=DD1-1 ; D21:=DD1 ; end;
end;

(*======================================================================*)
procedure IntervalleDateBalGL2(MonoExo : Boolean ; LExo : TExoDate ; Var D1,D11,D21 : TDateTime) ;
Var DD1 : TDateTime ;
    Exo : TExoDate ;
begin
DD1:=D1 ;
if MonoExo then Exo:=LExo else Exo:=VH^.EnCours ;
if Exo.Deb=D1 then begin D11:=D1 ; D21:=D11 ; end else begin D1:=Exo.Deb ; D11:=DD1-1 ; D21:=DD1 ; end;
end;

(*======================================================================*)
procedure InitTabColBal(Var TabCol : TTabCollAff) ;
Var  i :Integer ;
begin
Fillchar(TabCol,SizeOf(TabCol),#0) ;
For i:=1 To 6 Do  begin TabCol[i].OkAff:=FALSE ; TabCol[i].Tag[1]:=-1; TabCol[i].Tag[2]:=-1 ; end;
end;

(*======================================================================*)
{$IFNDEF EAGLCLIENT}
procedure SauvePositionTitre(CL : TQRLabel) ;
Var IsCritere : Boolean ;
    St1,St2 : String ;
begin
IsCritere:=CL.SynCritere>0 ;
if IsCritere And (CL.SynData='') then
   begin
   St1:=IntToStr(CL.Left) ; St2:=IntToStr(CL.Top) ; CL.SynData:=St1+';'+St2+';' ;
   end;
end;

(*======================================================================*)
procedure RestorePositionTitre(CL : TQRLabel ;var MaxTop : Integer) ;
Var IsCritere : Boolean ;
    St1,St : String ;
begin
IsCritere:=CL.SynCritere>0 ;
if IsCritere And (CL.SynData<>'') then
   begin
   St:=CL.SynData ;
   St1:=Trim(ReadTokenSt(St)) ; CL.Left:=StrToInt(st1) ;
   St1:=Trim(ReadTokenSt(St)) ; CL.Top:=StrToInt(st1) ;
   end;
if IsCritere And (MaxTop<CL.Top) then MaxTop:=CL.Top ;
end;

(*======================================================================*)
procedure InitFormatBal ( FF : TForm ; Var PremTabCol : TTabCollAff) ;
Type TTranslate = Array[1..6] Of Record Ori,Dest : Integer ; end;

Var i,j,k : integer ;
    T   : TComponent ;
    CL  : TQRLabel ;
begin
InitTabColBal(PremTabCol) ;
for i:=0 to FF.ComponentCount-1 do
    begin
    T:=FF.Components[i] ;
    if (T is TQRLabel) then
       begin
       CL:=TQRLabel(T) ; j:=Abs(CL.Tag) ;
       if CL.SynTitreCol And ((CL.SynColGroup=0) Or (CL.SynColGroup=100)) And (j>0) then
         Case j Of
           1,2,3 : PutTabCol(j,PremTabCol[j],CL.Tag,TRUE) ;
           5 : PutTabCol(4,PremTabCol[4],CL.Tag,TRUE) ;
           7 : PutTabCol(5,PremTabCol[5],CL.Tag,TRUE) ;
           9 : PutTabCol(6,PremTabCol[6],CL.Tag,TRUE) ;
           end;
       end;
    end;
for i:=0 to FF.ComponentCount-1 do
  begin
  T:=FF.Components[i] ;
  if (T is TQRLabel) then
     begin
     CL:=TQRLabel(T) ;
     (*
     For j:=1 To 6 Do For k:=1 To 2 Do
        if (Not CL.SynTitreCol) And (Abs(CL.tag)=PremTabCol[j].Tag[k]) then
           begin
           PremTabCol[j].Left[k]:=CL.Left ; PremTabCol[j].Width[k]:=CL.Width ;
           end;
     *)
     if CL.SynCritere>100 then SauvePositionTitre(CL) ;
     if (CL.SynTitreCol) And (CL.SynColGroup<>100) then
        begin
        For j:=1 To 6 Do For k:=1 To 2 Do
           begin
           if (Abs(CL.tag)=PremTabCol[j].Tag[k]) then
              begin
              PremTabCol[j].Left[k]:=CL.Left ; PremTabCol[j].Width[k]:=CL.Width ;
              end;
           end;
        end;
     end;
  end;
For j:=1 To 6 Do For k:=1 To 2 Do
//DebugV([j,k,Abs(CL.tag),PremTabCol[j].Left[k],PremTabCol[j].Width[k],'-----------']) ;
end;

(*======================================================================*)
function ChangeFormatBal ( FF : TForm ; Var PremTabCol,TabCol : TTabCollAff) : Integer ;
Type TTranslate = Array[1..6] Of Record Ori,Dest : Integer ; end;

Var i,j,k,ll,CLSynCOLGROUP : integer ;
    T   : TComponent ;
    C   : TQRCustomControl ;
    CL  : TQRLabel ;
    IsTitre : Boolean ;
    TT  : TTranslate ;
begin
For i:=1 To 6 Do if PremTabCol[i].OkAff then
   begin
{   TabCol[i].OkAff:=TRUE ; TabCol[i].Tag:=PremTabCol[i].Tag ;}
   end else
   begin
   TabCol[i].OkAff:=FALSE ; TabCol[i].Tag[1]:=-1 ; TabCol[i].Tag[2]:=-1 ;
   end;
for i:=0 to FF.ComponentCount-1 do
    begin
    T:=FF.Components[i] ;
    if (T is TQRLabel) Or (T Is TQRDBText) then
       begin
       IsTitre:=(T is TQRLabel) ; CLSynColGroup:=0 ;
       if ISTitre then begin CL:=TQRLabel(T) ; IsTitre:=CL.SynTitreCol ; CLSynColGroup:=CL.SynColGroup ; end;
       C:=TQRCustomControl(T) ;
       For j:=1 To 6 Do For k:=1 To 2 Do
         begin
         if Abs(C.tag)=TabCol[j].Tag[k] then
            begin
            if IsTitre And (CLSynColGroup=100) then { Titre avec sous colonne }
               begin
               C.Left:=PremTabCol[j].Left[1] ;
               C.Width:=PremTabCol[j].Width[1]*2+1 ;
               end else
               begin
               C.Left:=PremTabCol[j].Left[k] ;
               C.Width:=PremTabCol[j].Width[k] ;
               end;
            C.Visible:=TabCol[j].OkAff ;
            if Not C.Visible then C.Width:=0 ;
            end;
         end;
       end;
    end;
ll:=0 ; k:=0 ;
For j:=1 To 6 Do
   begin
   TT[j].Ori:=J ; TT[j].Dest:=J ;
   if TabCol[j].OkAff then begin Inc(k) ; TT[j].Dest:=k ; end;
   end;
for i:=0 to FF.ComponentCount-1 do
   begin
   T:=FF.Components[i] ;
   if (T is TQRLabel) Or (T Is TQRDBText) then
      begin
      C:=TQRCustomControl(T) ;
      IsTitre:=(T is TQRLabel) ; CL:=NIL ;
      if ISTitre then begin CL:=TQRLabel(T) ; IsTitre:=CL.SynTitreCol ; end;
      For j:=1 To 6 Do
        begin
        For k:=1 To 2 Do
           begin
           if Abs(C.tag)=TabCol[j].Tag[k] then
              begin
              if TabCol[j].OkAff then
                 begin
                 C.Left:=PremTabCol[tt[j].Dest].Left[k] ;
                 if CL<>NIL then
                    if (k=1) And (IsTitre And TabCol[j].OkAff And ((CL.SynColGroup=0) Or (CL.SynColGroup=100))) then
                       ll:=ll+C.Width+1 ;
                 end;
              end;
           end;
        end;
      end;
   end;
Result:=ll ;
end;
{$ENDIF}

(*======================================================================*)
procedure PutTabCol(i : Integer ; Var Coll : TCollAff ; Tag : Integer ; Affiche : Boolean) ;
begin
if i<3 then
   begin
   Coll.Tag[1]:=Abs(Tag) ; Coll.Tag[2]:=Abs(Tag) ; Coll.OkAff:=Affiche ;
   end else
   begin
   Coll.Tag[1]:=Abs(Tag) ; Coll.OkAff:=Affiche ; Coll.Tag[2]:=Abs(Tag)+1 ;
   end;
end;

(*======================================================================*)
{$IFNDEF EAGLCLIENT}
procedure ChangeFormatCompte ( fb1,fb2 : TFichierBase ; FF : TForm ; Lg1,Tag1,Tag2 : Integer ; NomCompte : String) ;
Var T   : TComponent ;
    C   : TQRCustomControl ;
    Decal,Lg,Plus,i,OldWidth,Decal1 : Integer ;
    St : String ;
    fb : TFichierBase ;
    Debut : Boolean ;

begin
fb:=fb1 ; if VH^.Cpta[fb1].Lg<VH^.Cpta[fb2].Lg then fb:=fb2 ;
Lg:=VH^.Cpta[fb].Lg ; Debut:=TRUE ; if Tag1>1 then Debut:=FALSE ;
if Lg<8 then Lg:=8 ;
Decal:=Round(Lg*Lg1/17)+1 ;
//T:=NIL ;
T:=FF.FindComponent(NomCompte) ;
if T<>NIL then
  begin
  C:=TQRCustomControl(T) ;
  if VH^.Specif[0] then St:=Copy('WAWAWAWAWAWAWAWAW',1,Lg)
                   else St:=Copy('AZERTYUIOPMLKJJHG',1,Lg) ;
  OldWidth:=C.Width ;
  C.Width:=C.Canvas.TextWidth(St)+6 ;
  Decal:=OldWidth-C.Width ;
  C.Width:=OldWidth ;
  end;
for i:=0 to FF.ComponentCount-1 do
    begin
    T:=FF.Components[i] ;
    if (T is TQRLabel) Or (T Is TQRDBText) then
       begin
       C:=TQRCustomControl(T) ; Plus:=-1 ;
       if (Abs(C.tag)=Tag1) Or (Abs(C.tag)=Tag2) then
          begin
          if Abs(C.tag)=Tag2 then
             begin
             Plus:=1 ; C.Left:=C.Left-Decal ; Decal1:=0 ;
             end else
             begin
             Decal1:=3 ; if Not debut then Decal1:=0 ;
             C.Left:=C.Left+Decal1 ;
             end;

          C.Width:=C.Width+Plus*Decal-Decal1 ;
          end;
       end;
    end;
end;
{$ENDIF}

(*======================================================================*)
function DEUX2UNCOMPTE(St1,St2 : String) : String ;
begin
if Trim(St2)='' then Result:=St1 else Result:=BourreLess(St1,fbGene)+St2 ;
end;

(*======================================================================*)
{$IFNDEF EAGLCLIENT}
procedure ALIMCRITAFF(Cl : TQRLABEL ; Var Tab : TabCritAff)  ;
Var i,Dispo,NoColCrit : Integer ;
begin
Dispo:=1 ; NoColCrit:=CL.SynCritere Div 100 ;
For i:=1 To 3 Do
  begin
  if Tab[i].ColCrit<>0 then Inc(Dispo) ;
  if Tab[i].ColCrit=NoColCrit then
     begin
     if Tab[i].DecalLeft>CL.Left then Tab[i].DecalLeft:=CL.Left ;
     if Tab[i].DecalTop<CL.Top then Tab[i].DecalTop:=CL.Top ;
     Exit ;
     end;
  end;
Tab[Dispo].ColCrit:=NoColCrit ; Tab[Dispo].DecalLeft:=CL.Left ;
Tab[Dispo].DecalTop:=CL.Top ;
end;
{$ENDIF}

(*======================================================================*)
function ADEPLACER(NoColCrit : Integer ; Tab : TabCritAff) : Integer ;
Var i : Integer ;
begin
Result:=0 ;
if NoColCrit>100 then NoColCrit:=NoColCrit Div 100 ;
For i:=1 To 3 Do if Tab[i].ColCrit=NoColCrit then begin Result:=i ; Exit ; end;
end;


(*======================================================================*)
function TROUVEDECALTOP(Tab : TabCritAff) : Integer ;
Var i : Integer ;
begin
Result:=0 ;
For i:=1 To 3 Do if Tab[i].ColCrit>0 then
   begin if Result<Tab[i].DecalTop then Result:=Tab[i].DecalTop ; end;
end;

(*======================================================================*)
{$IFNDEF EAGLCLIENT}
procedure RETAILLEPIEDPAGE(T : TComponent ; QRP : TQuickReport ; L : Integer) ;
Var CB  : TQrBand ;
    n : Integer ;
    TheName : String ;
begin
CB:=TQRBand(T) ;
For n:=0 To CB.ControlCount-1 Do
  begin
  TheName:=uppercase(CB.Controls[n].Name) ;
  if (TheName='RNUMVERSION') Or (TheName='QRLABEL34') then CB.Controls[n].Width:=QRP.QRPrinter.PageWidth-2 else
    if (TheName='RUTILISATEUR') Or (TheName='QRSYSDATA2') then
       begin
       CB.Controls[n].Left:=QRP.QRPrinter.PageWidth-CB.Controls[n].Width-2 ;
       end;
  end;
end;

(*======================================================================*)
procedure RETAILLEENTETEAUTREPAGE(T : TComponent ; QRP : TQuickReport ; L : Integer) ;
Var CB  : TQrBand ;
    n : Integer ;
    TheName : String ;
begin
CB:=TQRBand(T) ;
For n:=0 To CB.ControlCount-1 Do
  begin
  TheName:=uppercase(CB.Controls[n].Name) ;
  if (TheName='TITREBARRE') then CB.Controls[n].Width:=QRP.QRPrinter.PageWidth ;
  end;
end;

(*======================================================================*)
procedure CALCULPOURMISEENPAGE(L : Integer ; QRP : TQuickReport ; FF : TForm ; Var CFP : TFormatPrint) ;
Var Ratio : Double ;
    T   : TComponent ;
    C   : TQRCustomControl ;
    CL  : TQRLabel ;
    i,j,k,Prem,Dern,ii,LeftDebut,TopDebut,MaxTop   : Integer ;
    NumTrait : LongInt ;
    UnePasse,IsTitre,IsCritere,OkReduireCritere : Boolean ;
    Orientation : TPrinterOrientation ;
    TabCol : TTabCollAff ;
    TabCrit : TabCritAff ;

begin
Fillchar(CFP.Report.R,SizeOf(CFP.Report.R),#0) ; MaxTop:=89 ;
TabCol:=CFP.TabColl ; FillChar(TabCrit,SizeOf(TabCrit),#0) ; OkReduireCritere:=FALSE ;
Prem:=6 ; For i:=6 DownTo 1 Do if TabCol[i].OkAff then begin Prem:=i ; end;
Dern:=1 ; For i:=1 To 6 Do if TabCol[i].OkAff then begin Dern:=i ; end;
if L<=CMaxPortrait then Orientation:=poPortrait else Orientation:=poLandscape ;
QRP.Orientation:=Orientation ;
if L=0 then Exit ; QRP.PrepareForPrinter ;
Ratio:=QRP.QRPrinter.PageWidth/L ; LeftDebut:=1000 ; TopDebut:=0 ;
for i:=0 to FF.ComponentCount-1 do
    begin
    T:=FF.Components[i] ;
    if T is TQRBand And (uppercase(T.Name)='PIEDPAGE') then RetaillePiedPage(T,QRP,L) ;
    if T is TQRBand And (uppercase(T.Name)='ENTETEAUTREPAGE') then RetailleEnteteAutrePage(T,QRP,L) ;
    if (T is TQRLabel) Or (T is TQRDBTEXT) then
       begin
       C:=TQRCustomControl(T) ; CL:=NIL ;
       IsTitre:=(T is TQRLabel) ;
       IsCritere:=IsTitre ;
       if ISTitre then
          begin
          CL:=TQRLabel(T) ; IsTitre:=CL.SynTitreCol ; IsCritere:=CL.SynCritere>0 ;
          end;
       if IsCritere then RestorePositionTitre(CL,MaxTop) ;
       if IsCritere And (Orientation=poPortrait) And (CL<>NIL) then
          begin
          if CL.Left+CL.Width>CMaxPortrait then
             begin
             AlimCritAff(CL,TabCrit) ; OkReduireCritere:=TRUE ;
             end else
             begin
             if CL.Left<LeftDebut then LeftDebut:=CL.Left ;
             if CL.Top>TopDebut then TopDebut:=CL.Top ;
             end;
          end;
       For j:=1 To 6 Do For k:=1 To 2 Do
         begin
         if Abs(C.tag)=TabCol[j].Tag[k] then
            if TabCol[j].OkAff then
               begin
               if IsTitre then UnePasse:=(CL.SynColGroup=0) Or (CL.SynColGroup=100) Or (CL.SynColGroup=101) else UnePasse:=j<3 ;
               if (Not UnePasse) Or (k=1)  then
                  begin
                  C.Left:=trunc(C.Left*Ratio) ;
                  C.Width:=trunc(C.Width*Ratio) ;
                  if IsTitre And ((CL.SynColGroup=0) Or (CL.SynColGroup=100)) then
                     begin
                     TabCol[j].Left[k]:=C.Left ;
                     TabCol[j].Width[k]:=C.Width ;
                     end;
                  end;
               end;
         end;
       end;
    end;
for i:=0 to FF.ComponentCount-1 do
    begin
    T:=FF.Components[i] ;
    if T Is TQRBand then
      begin
      // Rony --> En Balance Composite Comptable, la zone est trop petite
      //          Tandis qu'en Edition par Cumul c'est OK
      //C:=TQRCustomControl(T) ; if C.Tag=100 then C.Height:=MaxTop+20 ;
      C:=TQRCustomControl(T) ; if C.Tag=100 then C.Height:=MaxTop+50 ;
      end;
    if T Is TQRLigne then
       begin
       C:=TQRCustomControl(T) ;
       NumTrait:=C.Tag ;
       if NumTrait=-10 then
          begin
          end else
          begin
          if NumTrait in [1..6] then
             begin
             C.Visible:=TabCol[NumTrait].OkAff ;
             C.Visible:=C.Visible And CFP.PRSepMontant ;
             if NumTrait=Dern then C.Visible:=TRUE ;
             C.Left:=TabCol[NumTrait].Left[1]+TabCol[NumTrait].Width[1]+1 ;
             end;
          if NumTrait=-1 then
             begin
             C.Visible:=TabCol[Prem{1}].OkAff ;
             C.Left:=TabCol[Prem].Left[1]
             end;
          end;
       end else if (Orientation=poPortrait) And OkReduireCritere then
       begin
       if (T Is TQRLabel) then
          begin
          CL:=TQRLabel(T) ; IsCritere:=CL.SynCritere>0 ;
          if IsCritere then
             begin
             ii:=ADeplacer(CL.SynCritere,TabCrit) ;
             if ii>0 then
                begin
                CL.Left:=LeftDebut+CL.Left-TabCrit[ii].DecalLeft ;
                Cl.Top:=TopDebut+15+Cl.Top-41 ;
                end;
             end;
          end else if T Is TQRBand then
          begin
          C:=TQRCustomControl(T) ;
          if C.Tag=100 then C.Height:=C.Height+TrouveDecalTop(TabCrit) ;
          end;
       end;
    end;
end;

(*======================================================================*)
procedure InitEdit(Entete : TQRBAND ; QRP : TQuickReport) ;
begin
Entete.Frame.DrawTop:=Not V_PGI.QRCouleur ;
Entete.Frame.DrawBottom:=Not V_PGI.QRCouleur ;
Entete.Frame.DrawLeft:=Not V_PGI.QRCouleur ;
Entete.Frame.DrawRight:=Not V_PGI.QRCouleur ;
QRP.LeftMargin:=0 ; QRP.TopMargin:=0 ; QRP.Columns:=1 ;
QRP.ColumnMargin:=0 ; QRP.RowMargin:=0 ;
QRP.QRPrinter.SynEdtCpta:=TRUE ;
end;

(*======================================================================*)
procedure ReInitEdit(Entete : TQRBAND ; AvecCouleur : Boolean) ;
begin
Entete.Frame.DrawTop:=Not AvecCouleur ;
Entete.Frame.DrawBottom:=Not AvecCouleur ;
Entete.Frame.DrawLeft:=Not AvecCouleur ;
Entete.Frame.DrawRight:=Not AvecCouleur ;
end;
{$ENDIF}

(*======================================================================*)
procedure InitTabCol(Var TabCol : TTabCollAffEdt ; OkAffiche : Boolean) ;
Var  i :Integer ;
begin
Fillchar(TabCol,SizeOf(TabCol),#0) ;
For i:=1 To MaxColEdt Do  begin TabCol[i].OkAff:=OkAffiche ; TabCol[i].Tag:=-1; end;
end;

(*======================================================================*)
{$IFNDEF EAGLCLIENT}
procedure InitFormatEdt ( FF : TForm ; Var PremTabCol,TabCol : TTabCollAffEdt) ;
Type TTranslate = Array[1..MaxTranslate] Of Record Ori,Dest : Integer ; end;

Var i,j : integer ;
    T   : TComponent ;
    CL  : TQRLabel ;
begin
InitTabCol(PremTabCol,TRUE) ;
for i:=0 to FF.ComponentCount-1 do
    begin
    T:=FF.Components[i] ;
    if (T is TQRLabel) then
       begin
       CL:=TQRLabel(T) ; j:=Abs(CL.Tag) ;
       if CL.SynTitreCol And ((CL.SynColGroup=0) Or (CL.SynColGroup=100)) And (j>0) then
          begin
          PremTabCol[j].Tag:=j ; PremTabCol[j].OkAff:=TRUE ;
          end;
       end;
    end;
for i:=0 to FF.ComponentCount-1 do
  begin
  T:=FF.Components[i] ;
  if (T is TQRLabel) then
     begin
     CL:=TQRLabel(T) ;
     For j:=1 To MaxColEdt Do
        if ({Not} CL.SynTitreCol) And (Abs(CL.tag)=PremTabCol[j].Tag) then
           begin
           PremTabCol[j].Left:=CL.Left ; PremTabCol[j].Width:=CL.Width ;
           end;
     end;
  end;
For i:=1 To MaxColEdt Do if PremTabCol[i].Tag=-1 then PremTabCol[i].OkAff:=FALSE ;
For i:=1 To MaxColEdt Do if PremTabCol[i].OkAff then
   begin
   TabCol[i].OkAff:=TRUE ; TabCol[i].Tag:=PremTabCol[i].Tag ;
   end;
end;
{$ENDIF}

(*======================================================================*)
function TrouveDecal(J : Integer ; Var PremTabCol,TabCol : TTabCollAffEdt) : Integer ;
Var i : Integer ;
    ll : Integer ;
begin
ll:=0 ;
For i:=j-1 DownTo 1 do begin
  if Not TabCol[i].OkAff then ll:=ll+PremTabCol[i].Width ;
end;
Result:=ll ;
end;

(*======================================================================*)
{$IFNDEF EAGLCLIENT}
function  ChangeFormatEdt ( FF : TForm ; Var PremTabCol,TabCol : TTabCollAffEdt ; initTabCol : Boolean = TRUE) : Integer ;
Type TTranslate = Array[1..MaxTranslate] Of Record Ori,Dest : Integer ; end;
Var i,j,k,ll,X,LeTag : integer ;
    T   : TComponent ;
    C   : TQRCustomControl ;
    CL  : TQRLabel ;
    IsTitre : Boolean ;
    TT  : TTranslate ;
begin
For i:=1 To MaxColEdt Do if PremTabCol[i].OkAff then
   begin
  // GC - 20/12/2001
   if InitTabCol then
     begin
   // GC - 20/12/2001 - FIN
     TabCol[i].OkAff:=TRUE ; TabCol[i].Tag:=PremTabCol[i].Tag ;
   // GC - 20/12/2001
     end;
   // GC - 20/12/2001 - FIN
   end else
   begin
   TabCol[i].OkAff:=FALSE ; TabCol[i].Tag:=-1 ;
   end;
for i:=0 to FF.ComponentCount-1 do
    begin
    T:=FF.Components[i] ;
    if (T is TQRLabel) Or (T Is TQRDBText) then
       begin
       IsTitre:=(T is TQRLabel) ;
       if ISTitre then
          begin
          CL:=TQRLabel(T) ; //IsTitre:=CL.SynTitreCol ;
          if Copy(CL.NAME,1,3)='SUP' then begin CL.SynData:='' ; CL.Caption:='' ; end;
          end;
       C:=TQRCustomControl(T) ;
       For j:=1 To MaxColEdt Do
         begin
         if Abs(C.tag)=TabCol[j].Tag then
            begin
            C.Left:=PremTabCol[j].Left ;
            C.Width:=PremTabCol[j].Width ;
            C.Visible:=TabCol[j].OkAff ;
            if Not C.Visible then C.Width:=0 ;
            end;
         end;
       end;
    end;
ll:=0 ; k:=0 ;
For j:=1 To MaxColEdt Do
   begin
   TT[j].Ori:=J ; TT[j].Dest:=J ;
   if TabCol[j].OkAff then begin Inc(k) ; TT[j].Dest:=k ; end;
   end;
for i:=0 to FF.ComponentCount-1 do
   begin
   T:=FF.Components[i] ;
   if (T is TQRLabel) Or (T Is TQRDBText) then
      begin
      C:=TQRCustomControl(T) ;
      IsTitre:=(T is TQRLabel) ;
      if ISTitre then begin CL:=TQRLabel(T) ; IsTitre:=CL.SynTitreCol ; end;
      For j:=1 To MaxColEdt Do
        begin
        if Abs(C.tag)=TabCol[j].Tag then
           begin
           if TabCol[j].OkAff then
              begin
              // GC - 20/12/2001
              X:=TrouveDecal(j,PremTabCol,TabCol) ;
//              C.Left:=PremTabCol[tt[j].Dest].Left-X ;
              C.Left:=PremTabCol[tt[j].Ori].Left-X ;
              // GC - 20/12/2001 - FIN
              if IsTitre And TabCol[j].OkAff then ll:=ll+C.Width+1 ;
              end;
           end;
        end;
      LeTag:=Abs(C.Tag) ;
      if (LeTag>0) And (LeTag<=MaxColEdt) then if Not TabCol[LeTag].OkAff then begin C.Visible:=FALSE ; C.Width:=0 ; end;
      end;
   end;
Result:=ll ;
end;

(*======================================================================*)
procedure CALCULPOURMISEENPAGEEDT(L : Integer ; QRP : TQuickReport ; FF : TForm ; Var CFP : TFormatPrintEdt) ;
Var Ratio : Double ;
    T   : TComponent ;
    C   : TQRCustomControl ;
    CL  : TQRLabel ;
    i,j,Prem,Dern,ii,LeftDebut,TopDebut  : Integer ;
    NumTrait : LongInt ;
    IsTitre,IsCritere,OkReduireCritere : Boolean ;
    Orientation : TPrinterOrientation ;
    TabCol : TTabCollAffEdt ;
    TabCrit : TabCritAff ;

begin
Fillchar(CFP.Report.R,SizeOf(CFP.Report.R),#0) ;
TabCol:=CFP.TabColl ; FillChar(TabCrit,SizeOf(TabCrit),#0) ; OkReduireCritere:=FALSE ;
Prem:=MaxColEdt ; For i:=MaxColEdt DownTo 1 Do if TabCol[i].OkAff then begin Prem:=i ; end;
Dern:=1 ; For i:=1 To MaxColEdt Do if TabCol[i].OkAff then begin Dern:=i ; end;
if CFP.ModePrint=PPortrait then
   begin
   Orientation:=poPortrait ;
   QRP.Orientation:=Orientation ;
   if L=0 then Exit ; QRP.PrepareForPrinter ;
   Ratio:=QRP.QRPrinter.PageWidth/L ; LeftDebut:=1000 ; TopDebut:=0 ;
   end else
   begin
   if L<=CMaxPortrait then Orientation:=poPortrait else Orientation:=poLandscape ;
   QRP.Orientation:=Orientation ;
   if L=0 then Exit ; QRP.PrepareForPrinter ;
   Ratio:=QRP.QRPrinter.PageWidth/L ; LeftDebut:=1000 ; TopDebut:=0 ;
   end;
for i:=0 to FF.ComponentCount-1 do
    begin
    T:=FF.Components[i] ;
    if T is TQRBand And (uppercase(T.Name)='PIEDPAGE') then RetaillePiedPage(T,QRP,L) ;
    if T is TQRBand And (uppercase(T.Name)='ENTETEAUTREPAGE') then RetailleEnteteAutrePage(T,QRP,L) ;
    if (T is TQRLabel) Or (T is TQRDBTEXT) then
       begin
       C:=TQRCustomControl(T) ;
       IsTitre:=(T is TQRLabel) ;
       IsCritere:=IsTitre ; CL:=NIL ;
       if ISTitre then
          begin
          CL:=TQRLabel(T) ; IsTitre:=CL.SynTitreCol ; IsCritere:=CL.SynCritere>0 ;
          end;
       if IsCritere And (Orientation=poPortrait) And (CL<>NIL) then
          begin
          if CL.Left+CL.Width>CMaxPortrait then begin AlimCritAff(CL,TabCrit) ; OkReduireCritere:=TRUE ; end else
             begin
             if CL.Left<LeftDebut then LeftDebut:=CL.Left ;
             if CL.Top>TopDebut then TopDebut:=CL.Top ;
             end;
          end;
       For j:=1 To MaxColEdt Do
         begin
         if Abs(C.tag)=TabCol[j].Tag then
            if TabCol[j].OkAff then
               begin
               C.Left:=trunc(C.Left*Ratio) ;
               C.Width:=trunc(C.Width*Ratio) ;
               if IsTitre And ((CL.SynColGroup=0) Or (CL.SynColGroup=100)) then
                  begin
                  TabCol[j].Left:=C.Left ;
                  TabCol[j].Width:=C.Width ;
                  end;
               end;
         end;
       end;
    end;
for i:=0 to FF.ComponentCount-1 do
    begin
    T:=FF.Components[i] ;
    if T Is TQRLigne then
       begin
       C:=TQRCustomControl(T) ;
       NumTrait:=C.Tag ;
       if NumTrait=-10 then
          begin
          end else
          begin
          if NumTrait in [1..MaxColEdt] then
             begin
             C.Visible:=TabCol[NumTrait].OkAff ;
             C.Visible:=C.Visible And CFP.PRSepMontant ;
             if NumTrait=Dern then C.Visible:=TRUE ;
             C.Left:=TabCol[NumTrait].Left+TabCol[NumTrait].Width+1 ;
             end;
          if NumTrait=-1 then
             begin
             C.Visible:=TabCol[Prem].OkAff ;
             C.Left:=TabCol[Prem].Left ;
             end;
          end;
       end else if (Orientation=poPortrait) And OkReduireCritere then
       begin
       if (T Is TQRLabel) then
          begin
          CL:=TQRLabel(T) ; IsCritere:=CL.SynCritere>0 ;
          if IsCritere then
             begin
             ii:=ADeplacer(CL.SynCritere,TabCrit) ;
             if ii>0 then
                begin
                CL.Left:=LeftDebut+CL.Left-TabCrit[ii].DecalLeft ;
                Cl.Top:=TopDebut+Cl.Top ;
                end;
             end;
          end else if T Is TQRBand then
          begin
          C:=TQRCustomControl(T) ; if C.Tag=100 then C.Height:=C.Height+TrouveDecalTop(TabCrit) ;
          end;
       end;
    end;
end;
{$ENDIF}

(*======================================================================*)
procedure ADDREPORT(Seti : SetOfByte ; Var Report : TabReport ; D,C : Double ;Decimale : Integer ) ;
Var i : Integer ;
begin
For i:=1 To 10 Do if i In SetI then
   begin
   Report.R[i][1].TotDebit:=Arrondi(Report.R[i][1].TotDebit+D,Decimale) ;
   Report.R[i][1].TotCredit:=Arrondi(Report.R[i][1].TotCredit+C,Decimale) ;
   Report.R[i][1].Active:=TRUE ;
   end;
end;

(*======================================================================*)
procedure INITREPORT(Seti : SetOfByte ; Var Report : TabReport) ;
Var i,j : Integer ;
begin
For i:=1 To 10 Do if i In SetI then For j:=1 To MaxNbRep Do
   begin
   Report.R[i][j].TotDebit:=0 ; Report.R[i][j].TotCredit:=0 ; Report.R[i][j].Active:=FALSE ;
   end;
end;

(*======================================================================*)
function QUELREPORT(Report : TabReport ; Var D,C : Double) : Byte ;
Var i : Integer ;
begin
D:=0 ; C:=0 ; Result:=0 ;
For i:=10 DownTo 1 Do
  if (Report.R[i][1].TotDebit<>0) Or (Report.R[i][1].TotCredit<>0) Or (Report.R[i][1].Active) then
     begin
     Result:=i ; D:=Report.R[i][1].TotDebit ; C:=Report.R[i][1].TotCredit ; Exit ;
     end;
end;

(*======================================================================*)
procedure ADDREPORTBAL(Seti : SetOfByte ; Var Report : TabReport ; Tot : TabTRep ;Decimale : Integer ) ;
Var i,j : Integer ;
begin
For i:=1 To 10 Do if i In SetI then For j:=1 To MaxNbRep Do
   begin
   Report.R[i][j].TotDebit:=Arrondi(Report.R[i][j].TotDebit+Tot[j].TotDebit,Decimale) ;
   Report.R[i][j].TotCredit:=Arrondi(Report.R[i][j].TotCredit+Tot[j].TotCredit,Decimale) ;
   Report.R[i][j].Active:=TRUE ;
   end;
end;

(*======================================================================*)
function Okrep(Rep : TReport) : Boolean ;
begin
Result:=(Rep.TotDebit<>0) Or (Rep.TotCredit<>0);
end;

(*======================================================================*)
function QUELREPORTBAL(Report : TabReport ; Var Tot : TabTRep) : Byte ;
Var i,j,k : Integer ;
    OkOk : Boolean ;
begin
FillChar(Tot,SizeOf(Tot),#0) ; Result:=0 ; OkOk:=FALSE ;
For i:=10 DownTo 1 Do
  begin
  For k:=1 To MaxNbRep Do OkOk:=OkOk Or OkRep(Report.R[i][k]) ;
  OkOk:=(Report.R[i][1].TotDebit<>0) Or (Report.R[i][1].TotCredit<>0) Or
        (Report.R[i][2].TotDebit<>0) Or (Report.R[i][2].TotCredit<>0) Or
        (Report.R[i][3].TotDebit<>0) Or (Report.R[i][3].TotCredit<>0) Or
        (Report.R[i][4].TotDebit<>0) Or (Report.R[i][4].TotCredit<>0) Or
        (Report.R[i][5].TotDebit<>0) Or (Report.R[i][5].TotCredit<>0) Or
        (Report.R[i][1].Active) ;
  if OkOk then
     begin
     Result:=i ;
     For j:=1 To MaxNbRep Do
        begin
        Tot[j].TotDebit:=Report.R[i][j].TotDebit ;
        Tot[j].TotCredit:=Report.R[i][j].TotCredit ;
        end;
     Exit ;
     end;
  end;
end;

(*======================================================================*)
procedure POSITIONNEFOURCHETTEST(TC1,TC2 : TComponent ; Var ST1,ST2 : String) ;
Var St : String ;
    Q : TQuery ;
    HCpt1,HCpt2 : THCpteEdit ;
begin
ST1:='' ; ST2:='' ;
if (TC1 Is THCpteEdit) And (TC2 Is THCpteEdit) then
   begin
   HCpt1:=THCpteEdit(TC1) ; HCpt2:=THCpteEdit(TC2) ;
   ST1:=HCpt1.Text ; ST2:=HCpt2.Text ;
   if (HCpt1.Text='') And (HCpt2.Text='') then
      begin
      Case CaseFic(HCpt1.ZoomTable) Of
        fbGene : St:='SELECT MIN(G_GENERAL), Max(G_GENERAL) FROM GENERAUX WHERE G_FERME="-" ' ;
        fbAux : St:='SELECT MIN(T_AUXILIAIRE), Max(T_AUXILIAIRE) FROM TIERS WHERE T_FERME="-" ' ;
        fbJal : St:='SELECT MIN(J_JOURNAL), Max(J_JOURNAL) FROM JOURNAL WHERE J_FERME="-" ' ;
        fbAxe1..fbAxe5 : St:='SELECT MIN(S_SECTION), Max(S_SECTION) FROM SECTION WHERE S_FERME="-" ' ;
        fbBudGen : St:='SELECT MIN(BG_BUDGENE), Max(BG_BUDGENE) FROM BUDGENE WHERE BG_FERME="-" ' ;
        fbBudJal : St:='SELECT MIN(BJ_BUDJAL), Max(BJ_BUDJAL) FROM BUDJAL WHERE BJ_FERME="-" ' ;
        fbBudSec1..fbBudSec5 : St:='SELECT MIN(BS_BUDSECT), Max(BS_BUDSECT) FROM BUDSECT WHERE BS_FERME="-" ' ;
        fbNatCpt : St:='SELECT MIN(NT_NATURE), Max(NT_NATURE) FROM NATCPTE WHERE NT_SOMMEIL="-" ' ;
        end;
      St:=St+RecupWhere(HCpt1.ZoomTable) ;
      Q:=OpenSQL(St,TRUE) ;
      if Not Q.EOF then begin St1:=Q.Fields[0].AsString ; St2:=Q.Fields[1].AsString ; end;
      Ferme(Q);
      end;
   end else
   begin
   Exit ;
   end;
end;

(*======================================================================*)
{$IFNDEF EAGLCLIENT}
procedure POSITIONNEFOURCHETTE(TC1,TC2 : TComponent ; QRL1,QRL2 : TQRLABEL) ;
Var St1,St2 : String ;
begin
PositionneFourchetteSt(TC1,TC2,St1,St2) ;
QRL1.Caption:=St1 ; QRL2.Caption:=St2 ;
end;
{$ENDIF}

(*======================================================================*)
function LeTypeField(LLFD : TStringList ; Nom : String ; Var LaSize : Word) : TFieldType ;
Var i : Integer ;
    St : String ;
    TypeField : TFieldType ;
begin
Result:=ftUnknown ; LaSize:=0 ;
For i:=0 To LLFD.Count-1 Do
  begin
  (*
  if Q.Fielddefs.Items[i].Name=Nom then
     begin
     Result:=Q.Fielddefs.Items[i].DataType ; Exit ;
     end;
  *)
  TypeField:=TFieldType(StrToInt(Copy(LLFD.Strings[i],1,2))) ;
  LaSize:=StrToInt(Copy(LLFD.Strings[i],3,3)) ;
  St:=Copy(LLFD.Strings[i],6,Length(LLFD.Strings[i])-5) ;
  if St=Nom then begin Result:=TypeField ; Exit ; end;
  end;
end;

(*======================================================================*)
function IsTField(LLF : TStringList ; Nom : String) : Boolean ;
Var i : Integer ;
begin
Result:=FALSE ;
For i:=0 To LLF.Count-1 Do
  begin
  {
  if Q.Fields[i].FieldName=Nom then begin Result:=TRUE ; Exit ; end;
  }
  if LLF.Strings[i]=Nom then begin Result:=TRUE ; Exit ; end;
  end;
end;

(*======================================================================*)
{$IFNDEF EAGLCLIENT}
procedure AddTfield(FF : TForm ; Q : TQuery ; LLFD : TStringList ; Nom : String) ;
Var TF : TField ;
    TypeField : TFieldType ;
    laSize : Word ;
begin
TypeField:=LeTypeField(LLFD,Nom,LaSize) ;
if TypeField=ftUnknown then Exit ;
Case TypeField Of
  ftString  : TF:=tStringField.Create(FF) ;
  ftInteger : TF:=TIntegerField.Create(FF) ;
  ftSmallInt : TF:=TSmallIntField.Create(FF) ;
  ftDateTime : TF:=TDateTimeField.Create(FF) ;
  ftFloat : TF:=TFloatField.Create(FF) ;
  else exit ;
  end;
TF.Name:=Q.Name+Nom ;
TF.FieldName:=Nom ;
Case TF.DataType Of
  ftFloat : TFloatField(TF).DisplayFormat:='#,##.0' ;
  ftString : TStringField(TF).Size:=LaSize ;
  end;
TF.Dataset:=Q ;
TF.Tag:=123 ;
end;
{$ENDIF}

(*======================================================================*)
function READTOKENST1 ( Var St : String) : String ;
Var i : Integer ;
begin
i:=Pos(',',St) ; if i<=0 then i:=Length(St)+1 ; Result:=Copy(St,1,i-1) ; Delete(St,1,i) ;
end;

(*======================================================================*)
{$IFNDEF EAGLCLIENT}
procedure CREATETFIELD(FF : TForm ; Q : TQuery ; SQLGA : String ; LLF,LLFD : TStringList) ;
Var St,St1 : String ;
begin
St:=SQLGA ;
While St<>'' Do
  begin
  St1:=Trim(ReadTokenSt1(St)) ;
  if Not IsTField(LLF,St1) then AddTField(FF,Q,LLFD,St1) ;
  end;
end;

(*======================================================================*)
procedure PREPARECREATETFIELD( Q : TQuery ; Var LLF,LLFD : TStringList) ;
Var i : Integer ;
begin
LLF:=TStringList.Create ; LLF.Sorted:=TRUE ; LLF.Duplicates:=DupIgnore ;
LLFD:=TStringList.Create ; LLFD.Sorted:=TRUE ; LLFD.Duplicates:=DupIgnore ;
For i:=0 To Q.Fielddefs.Count-1 Do
   begin
   LLFD.Add(FormatFloat('00',Ord(Q.Fielddefs.Items[i].DataType))+
            FormatFloat('000',Ord(Q.Fielddefs.Items[i].Size))+
            Q.Fielddefs.Items[i].Name) ;
   end;
For i:=0 To Q.FieldCount-1 Do LLF.Add(Q.Fields[i].FieldName) ;
end;

(*======================================================================*)
function GETPARAMLISTE(NewValue : String ; FF : TForm) : String;
Var Table,Crit,tri,Champ,Larg,Align,FPerso,Params : String ;
    Titre, LeTitre, NumCol : HString;
    StF,StL,StJ,StT,F,J,TI,NC,StNC : String ;
    i : Integer ;
    C : TQRLABEL ;
    Name : String ;
    T : TComponent ;
    OkTri,OkNumCol : Boolean ;
begin
Result:='' ; OkTri:=False ; OkNumCol:=False ;
if ChargeHListe(NewValue,Table,Crit,tri,Champ,Titre,Larg,Align,Params,LeTitre,NumCol,FPerso,OkTri,OkNumCol) then
   begin
   StF:=Champ ; While Pos(',',StF)>0 do StF[Pos(',',StF)]:=';' ;
   StL:=Larg ; StJ:=Align ; StT:=Titre ; StNC:=NumCol ;
//   i:=0 ;
   While StF<>'' do
      begin
      NC:=Trim(ReadTokenSt(StNC)) ;i:=StrToInt(NC) ;
      Name:='SUP'+IntToStr(i) ; T:=FF.Findcomponent(Name) ;
      if T<>NIL then
         begin
         F:=Trim(ReadTokenSt(StF)) ; J:=Trim(ReadTokenSt(StJ)) ; TI:=Trim(ReadTokenSt(StT)) ;
         if T<>NIL then
            begin
            C:=TQRLABEL(T) ;
            C.Visible:=TRUE ;
            if J[1]='D' then C.Alignment:=taRightJustify else
               if J[1]='G' then C.Alignment:=taLeftJustify else
                 if J[1]='C' then C.Alignment:=taCenter ;
            C.SynData:=C.SynData+F+';' ;
            if Pos('$',F)>0 then System.Delete(F,1,1) ;
            Result:=Result+F+', ' ;
            (*
            debugV1(['get :',Result]) ;
            if Not IsTField(Q,F) then AddTField(Q,F) ;
            *)
            end;
         end;
      end;

   end;
end;
{$ENDIF}

(*======================================================================*)
(*
procedure INITFIELDS(FF : TForm ; Q : TQuery ; SQLGA : String ; LLF,LLFD : TStringList) ;
Var St,St1 : String ;
begin
St:=SQLGA ;
While St<>'' Do
  begin
  St1:=Trim(ReadTokenSt1(St)) ;
  if Not IsTField(LLF,St1) then AddTField(FF,Q,LLFD,St1) ;
  end;
end;
*)

(*======================================================================*)
{$IFNDEF EAGLCLIENT}
procedure ALIMLISTE(Q : TQuery ; i : Integer ; FF : TForm) ;
Var C : TQRLABEL ;
    Name : String ;
    T : TComponent ;
    TF : TField ;
    St,St1,StF : String ;
begin
Name:='SUP'+IntToStr(i) ; T:=FF.Findcomponent(Name) ;
if T<>NIL then
   begin
   C:=TQrLabel(T) ;
   if C.Visible then
      begin
      C.Caption:='' ;
      StF:=C.SynData ;
      While STF<>'' Do
         begin
         St1:=Trim(ReadTokenSt(StF)) ;
         if Pos('$',St1)>0 then
            begin
            System.Delete(St1,1,1) ;
            (*
            Ou:=CorrespType(CorrespChamp(Copy(St1,Pos('_',St1)+1,30)),StCode,StWhere) ;
            CorrespBase(Ou,StTable,StPrefixe) ;
            nn:=StPrefixe+intToStr(nbS) ;
            F1:=StTable+' '+nn ;
            nn:=nn+'.' ;
            StF:=nn+StPrefixe+'_LIBELLE' ;
            *)
            end;
         TF:=Q.FindField(St1) ; St:='' ;
         if TF<>NIL then
           Case TF.DataType Of
             ftString : St:=TF.AsString ;
             ftSmallint,ftInteger,ftWord  : St:=IntToStr(TF.AsInteger);
             ftFloat,ftCurrency,ftBCD : St:=FormatFloat('',TF.AsFloat);
             ftDate,ftDateTime : St:=DateToStr(TF.AsDateTime) ;
             end;
         C.Caption:=C.Caption+St+' ' ;
         end;
      end else C.Caption:='' ;
   end;
end;

(*======================================================================*)
procedure InitQRPRINTER(QR : TQuickReport ; SurGrid,Reduire,Couleur : Boolean) ;
begin
QR.QRPrinter.SynOnGrid:=SurGrid ; QR.QRPrinter.SynCouleur:=Couleur ;
if Reduire then QR.QRPrinter.Thumbs:=2 else QR.QRPrinter.Thumbs:=1 ;
end;
{$ENDIF}

(*======================================================================*)
function CtrlPerExo(Date1,Date2 : TDateTime) : Boolean ;
{Contrôle les Dates des 2 Périodes }
begin
// Rony 10/04/97 Result:=Date1<=Date2 ;
Result:=(IsValidDate(DateToStr(Date1))and IsValidDate(DateToStr(Date2))and(Date1<=Date2)) ;
end;

(*======================================================================*)
function CtrlMouvementeSur(MouvementeSur : integer ; NatureEcr : String) : Boolean ;
{Contrôle que "Sur comptes non soldés" en soit pas avec "Situation" ou "prévision"   }
begin
Result:=TRUE ;
if ((NatureEcr='SSI') or (NatureEcr='PRE')) And (MouvementeSur=2) then Result:=FALSE ;
end;

(*======================================================================*)
procedure InitCritEdt (Var CritEdt : TcritEdt ) ;
begin
With CritEdt Do
  begin
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ; DeviseSelect:='' ;
  DeviseAffichee:=V_PGI.DevisePivot ; AfficheSymbole:=False ;
  MonoExo:=True ; Valide:='g' ; RuptTabLibre:=False ;

  Case NatureEtat Of
    neBal : begin
            Bal.AnalPur:='g' ; Bal.TypCpt:=1 ;
            QuelExoDate(Date1,Date2,MonoExo,Exo) ;
            if Not MonoExo then Bal.TypCpt:=3 ; // Périodes
            IntervalleDateBALGL1(MonoExo,Exo,Date1,Bal.Date11,Bal.Date21) ;
            With Bal.FormatPrint Do
              begin
              PrSepMontant:=True ;
              PrSepCompte:=True ;
              PrSepRupt:=True ;
              PutTabCol(1, TabColl[1],-1,True) ;
              PutTabCol(2, TabColl[2], 2, True) ;
              PutTabCol(3, TabColl[3], 3, True) ;
              PutTabCol(4, TabColl[4], 5, True) ;
              PutTabCol(5, TabColl[5], 7, True) ;
              PutTabCol(6, TabColl[6], 9, True) ;
              Report.OkAff:=True ;
              end;
            end;
    neGL  : begin
            Gl.AnalPur:='g' ; Gl.TypCpt:=1 ;
            Gl.NumPiece1:=0 ;  Gl.NumPiece2:=99999999 ;
            QuelExoDate(Date1,Date2,MonoExo,Exo) ;
            if Not MonoExo then GL.TypCpt:=3 ; // Périodes
            IntervalleDateBALGL1(MonoExo,Exo,Date1,GL.Date11,GL.Date21) ;
            With Gl.FormatPrint Do
              begin
              PrSepMontant:=True ;
              PrSepCompte[1]:=True ;
              PrSepCompte[2]:=True ;
              Report.OkAff:=True ;
              end;
            end;
    neJU  : begin
            JU.Sens:=2 ;
            With JU.FormatPrint Do
                 begin
                 PrSepMontant:=True ;
                 PrSepCompte[1]:=True ;
                 PrSepCompte[2]:=True ;
                 PrSepCompte[3]:=True ;
                 PrSepCompte[4]:=True ;
                 Report.OkAff:=True ;
                 end;
            end;
    neBro : begin
            end;
    neEch : begin
            Ech.ModePaie:='' ;
            // Echeancier:=FEcheancier.ItemIndex ;
            Ech.Sens:=2 ;
            QuelExoDate(Date1,Date2,MonoExo,Exo) ;
            (*
            if DansExo(V_PGI.Precedent,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Precedent ; end else
               if DansExo(V_PGI.EnCours,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.EnCours ; end else
                  if DansExo(V_PGI.Suivant,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Suivant ; end;
            *)
            With Ech.FormatPrint Do
                 begin
                 PrSepMontant:=True  ;
                 PrSepCompte[1]:=True ;
                 PrSepCompte[2]:=False ;
                 Report.OkAff:=True ;
                 end;
            end;
    neGlV : begin
            GlV.Ecart:=30 ;
            GlV.Sens:=2 ;
            With GlV.FormatPrint Do
              begin
              PrSepMontant:=True ;
              PrSepCompte[1]:=True ;
              Report.OkAff:=True ;
              end;
            end;
    neJal : begin
            Jal.NumPiece1:=0   ; Jal.NumPiece2:=99999999 ;
            QuelExoDate(Date1,Date2,MonoExo,Exo) ;
            (*
            if DansExo(V_PGI.Precedent,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Precedent ; end else
               if DansExo(V_PGI.EnCours,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.EnCours ; end else
                  if DansExo(V_PGI.Suivant,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Suivant ; end;
            *)
            With Jal.FormatPrint Do
                 begin
                 PrSepMontant:=True ;
                 PrSepCompte[1]:=True ;
                 PrSepCompte[2]:=True ;
                 PrSepCompte[3]:=True ;
                 PrSepCompte[4]:=True ;
                 PrSepCompte[5]:=True ;
                 PrSepCompte[6]:=True ;
                 Report.OkAff:=True ;
                 end;
            end;
(*
    neJaG : begin
            {
            if DansExo(V_PGI.Precedent,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Precedent ; end else
               if DansExo(V_PGI.EnCours,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.EnCours ; end else
                  if DansExo(V_PGI.Suivant,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Suivant ; end;
            }
            QuelExoDate(Date1,Date2,MonoExo,Exo) ;
            With JaG.FormatPrint Do
                 begin
                 PrSepMontant:=True ; PrSepCompte:=True ;
                 PutTabCol(1,TabColl[1],-1,True) ;
                 PutTabCol(2,TabColl[2],2,True) ;
                 PutTabCol(3,TabColl[3],3,True) ;
                 PutTabCol(4,TabColl[4],5,False) ;
                 end;
            end;
*)
    neCum : begin
            NatureCpt:='' ;  Etab:='' ;
            QuelExoDate(Date1,Date2,MonoExo,Exo) ;
            if Not MonoExo then Cum.TypCpt:=3 ; // Périodes
            IntervalleDateBALGL1(MonoExo,Exo,Date1,Cum.Date11,Cum.Date21) ;
            Cum.AvecDevise:=False ;
            Cum.TypCpt:=1 ;
            With Cum.FormatPrint Do
                 begin
                 PrSepMontant:=True ; PrSepCompte:=True ;
                 PutTabCol(1, TabColl[1],-1, True) ;
                 PutTabCol(2, TabColl[2], 2, True) ;
                 PutTabCol(3, TabColl[3], 3, True) ;
                 PutTabCol(4, TabColl[4], 5, True) ;
                 PutTabCol(5, TabColl[5], 7, False) ;
                 PutTabCol(6, TabColl[6], 9, False) ;
                 end;
            end;
    nePoi : begin
            With Poi.FormatPrint Do
              begin
              prSepMontant:=TRUE ;
              PrSepCompte[1]:=TRUE ;
              PrSepCompte[2]:=TRUE ;
              PrSepCompte[3]:=FALSE ;
              end;
            Poi.Banque:='' ; Poi.RefP1:='' ; Poi.RefP2:='' ;
            end;
    neRap : begin
            QuelExoDate(Date2,Date2,MonoExo,Exo) ;
            DateDeb:=Exo.Deb ; Date1:=Exo.Deb ;
            DateFin:=DAte2 ;
            With Poi.FormatPrint Do
              begin
              prSepMontant:=TRUE ;
              end;
            Rap.RefP:='' ; Rap.Sens:=0 ;
            end;
     end;
  end;
end;

(*======================================================================*)
(*
procedure TFQR.RecupCritBal ;
Var ST : String ;
begin
Fillchar(CritEDT,SizeOf(CritEDT),#0) ;
With CritEDT Do
  begin
  NatureEtat:=neBal ;
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ;
  DeviseSelect:='' ; if FDevises.ItemIndex<>0 then DeviseSelect:=FDevises.Value ;
  DeviseAffichee:=V_PGI.DevisePivot ; Bal.CodeRupt1:='' ; Bal.CodeRupt2:='' ;
  Monnaie:=FMontant.ItemIndex ; if Monnaie=1 then DeviseAffichee:=DeviseSelect ;
  Joker:=FJoker.Visible ;
  if Joker then
     begin
     Cpt1:=FJoker.Text ; Cpt2:=FJoker.Text ;
     end else
     begin
     Cpt1:=FCpte1.Text ; Cpt2:=FCpte2.Text ;
     end;
  Date1:=StrToDate(FDateCompta1.Text) ; Date2:=StrToDate(FDateCompta2.Text) ;
  DateDeb:=Date1 ; DateFin:=Date2 ;
  if V_PGI.EnCours.Deb=Date1 then
     begin
     Bal.Date11:=StrToDate(FDateCompta1.Text) ; Bal.Date21:=Bal.Date11 ;
     end else
     begin
     Date1:=V_PGI.EnCours.Deb ; Bal.Date11:=StrToDate(FDateCompta1.Text)-1 ; Bal.Date21:=BAL.Date11+1 ;
     end;
  if FEtab.ItemIndex<>0 then Etab:=FEtab.Value ;
  QualifPiece:=FNatureEcr.Value ;
  if FNatureCpt.ItemIndex<>0 then NatureCpt:=FNatureCpt.Value ;
  Bal.TypCpt:=FSelectCpte.ItemIndex ;
  Bal.Sauf:=Trim(FExcep.Text) ;
//Simon
  if Bal.Sauf<>'' then
    begin
    Case FNatureBase of
       nbGen,nbGenT : if SqlCptInterdit('G_GENERAL', ST, FExcep) then ;
       nbAux : if SqlCptInterdit('T_AUXILIAIRE', ST, FExcep) then ;
       nbSec : if SqlCptInterdit('S_SECTION', ST, FExcep) then ;
       nbBud : if SqlCptInterdit('B_BUDGET', ST, FExcep) then ;
       end;
    Bal.SQLSauf:=St ;
    end;
  MonoExo:=FExercice.ItemIndex>0 ;
  if MonoExo then Exo.Code:=FExercice.Value ;
  if Not MonoExo then
     begin
     if DansExo(V_PGI.Precedent,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Precedent ; end else
        if DansExo(V_PGI.EnCours,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.EnCours ; end else
            if DansExo(V_PGI.Suivant,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Suivant ; end;
     if Not MonoExo then BAL.TypCpt:=3 ; // Périodes
     end;
  AfficheSymbole:=FMonetaire.Checked ;
  With Bal.FormatPrint Do
    begin
    PrSepMontant:=FTrait.Checked ;
    PrSepCompte:=FLigneCpt.Checked ;
    Report.OkAff:=FReport.Checked ;
    end;
  end;
end;
*)

(*======================================================================*)
(*
procedure TFQR.RecupCritGL ;
Var ST : String ;
begin
Fillchar(CritEDT,SizeOf(CritEDT),#0) ;
With CritEDT Do
  begin
  NatureEtat:=neGL ;
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ;
  DeviseSelect:='' ; if FDevises.ItemIndex<>0 then DeviseSelect:=FDevises.Value ;
  DeviseAffichee:=V_PGI.DevisePivot ; GL.CodeRupt1:='' ; GL.CodeRupt2:='' ;
  Monnaie:=FMontant.ItemIndex ; if Monnaie=1 then DeviseAffichee:=DeviseSelect ;
  Joker:=FJoker.Visible ;
  if Joker then
     begin
     Cpt1:=FJoker.Text ; Cpt2:=FJoker.Text ;
     end else
     begin
     Cpt1:=FCpte1.Text ; Cpt2:=FCpte2.Text ;
     end;
  Date1:=StrToDate(FDateCompta1.Text)    ; Date2:=StrToDate(FDateCompta2.Text) ;
  DateDeb:=DAte1 ; DateFin:=DAte2 ;
  if V_PGI.EnCours.Deb=Date1 then
     begin
     GL.Date11:=StrToDate(FDateCompta1.Text) ; GL.Date21:=GL.Date11 ;
     end else
     begin
     Date1:=V_PGI.EnCours.Deb ; GL.Date11:=DateDeb-1 ; GL.Date21:=GL.Date11+1 ;
     end;
  if FEtab.ItemIndex<>0 then Etab:=FEtab.Value ;
  if FNatureCpt.ItemIndex<>0 then NatureCpt:=FNatureCpt.Value ;
  GL.TypCpt:=FSelectCpte.ItemIndex ;
  GL.Sauf:=Trim(FExcep.Text) ;
//Simon
  if GL.Sauf<>'' then
    begin
    Case FNatureBase of
       nbGen,nbGenT : if SqlCptInterdit('G_GENERAL', St, FExcep) then ;
       nbAux : if SqlCptInterdit('T_AUXILIAIRE', St, FExcep) then ;
       nbSec : if SqlCptInterdit('S_SECTION', St, FExcep) then ;
       nbBud : if SqlCptInterdit('B_BUDGET', St, FExcep) then ;
       end;
    GL.SQLSauf:=St ;
    end;
  MonoExo:=FExercice.ItemIndex>0 ;
  if MonoExo then Exo.Code:=FExercice.Value ;
  if Not MonoExo then
     begin
     if DansExo(V_PGI.Precedent,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Precedent ; end else
        if DansExo(V_PGI.EnCours,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.EnCours ; end else
           if DansExo(V_PGI.Suivant,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Suivant ; end;
     if Not MonoExo then GL.TypCpt:=3 ; // Périodes
     end;
  AfficheSymbole:=FMonetaire.Checked ;
  With GL.FormatPrint Do
     begin
     PrSepMontant:=FTrait.Checked ;
     PrSepCompte[1]:=FLigneCpt.Checked ;
     Report.OkAff:=FReport.Checked ;
     end;
 end;
end;
*)

(*======================================================================*)
(*
procedure TFQR.RecupCritJu ;
begin
Fillchar(CritEDT,SizeOf(CritEDT),#0) ;
With CritEDT Do
  begin
  NatureEtat:=neJU ;
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ; DeviseSelect:='' ;
  DeviseAffichee:=V_PGI.DevisePivot ;
  Joker:=FJoker.Visible ;
  if Joker  then begin Cpt1:=FJoker.Text   ; Cpt2:=FJoker.Text ;  end
            else begin Cpt1:=Fcpte1.Text       ; Cpt2:=FCpte2.Text ;      end;
  Date1:=StrToDate(FDateCompta1.Text) ;
  if FDevises.ItemIndex<>0 then DeviseSelect:=FDevises.Value ;
  Monnaie:=FMontant.ItemIndex ;
  if Monnaie=1 then DeviseAffichee:=DeviseSelect ;
  if FEtab.ItemIndex<>0 then Etab:=FEtab.Value ;
  if FNatureCpt.ItemIndex<>0 then NatureCpt:=FNatureCpt.Value ;
  QualifPiece:=FNatureEcr.Value ;
  AfficheSymbole:=FMonetaire.Checked ;
  With JU.FormatPrint Do
       begin
       PrSepMontant:=FTrait.Checked ;
       PrSepCompte[1]:=FLigneCpt.Checked ;
       Report.OkAff:=FReport.Checked ;
       end;
  end;
end;
*)

(*======================================================================*)
{ -------------- Les Codes suivants sont provisoires -------------- }
{                       Mais ne pas ENLEVER !!!             ...Rony }
(*
procedure INITCRITJOUR (Var CritJour : TCritEdt ) ;
begin
With CritJour Do
  begin
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ;
  DeviseSelect:='' ; DeviseAffichee:=V_PGI.DevisePivot ;
  Valide:='g' ; MonoExo:=TRUE ; AfficheSymbole:=FALSE;

  Jal.NumPiece1:=0   ; Jal.NumPiece2:=99999999 ;
  if DansExo(V_PGI.Precedent,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Precedent ; end else
     if DansExo(V_PGI.EnCours,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.EnCours ; end else
        if DansExo(V_PGI.Suivant,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Suivant ; end;
  With Jal.FormatPrint Do
       begin
       PrSepMontant:=TRUE ;
       PrSepCompte[1]:=TRUE ;
       PrSepCompte[2]:=TRUE ;
       PrSepCompte[3]:=TRUE ;
       PrSepCompte[4]:=TRUE ;
       PrSepCompte[5]:=TRUE ;
       PrSepCompte[6]:=TRUE ;
       end;
  end;
end;
*)

(*======================================================================*)
(*
procedure INITCRITJOURGENE (Var CritJour : TCritEdt ) ;
begin
With CritJour Do
  begin
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ;
  DeviseAffichee:=V_PGI.DevisePivot ; DeviseSelect:='' ;
  AfficheSymbole:=False ; Valide:='g' ; MonoExo:=True ;

  if DansExo(V_PGI.Precedent,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Precedent ; end else
     if DansExo(V_PGI.EnCours,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.EnCours ; end else
         if DansExo(V_PGI.Suivant,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Suivant ; end;
  With JaG.FormatPrint Do
       begin
       PrSepMontant:=True ;  PrSepCompte:=True ;
       PutTabCol(1,TabColl[1],-1,True) ;
       PutTabCol(2,TabColl[2],2,True) ;
       PutTabCol(3,TabColl[3],3,True) ;
       PutTabCol(4,TabColl[4],5,False) ;
       end;
  end;
end;
*)

(*======================================================================*)
(*
procedure INITCRITCUMUL (Var CritCumul : TCritEdt ) ;
begin
With CritCumul Do
  begin
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ;
  DeviseAffichee:=V_PGI.DevisePivot ; DeviseSelect:='' ;
  Valide:='g' ; AfficheSymbole:=False; MonoExo:=True ;

  NatureCpt:='' ;  Etab:='' ;
  if V_PGI.EnCours.Deb=Date1
     then begin
          Cum.Date11:=Date1 ;
          Cum.Date21:=Cum.Date11 ;
          end
     else begin
          Date1:=V_PGI.EnCours.Deb ;
          Cum.Date11:=DateDeb-1 ;
          Cum.Date21:=Cum.Date11+1 ;
          end;
  Cum.AvecDevise:=False ;
  Cum.TypCpt:=1 ;
  if DansExo(V_PGI.Precedent,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Precedent ; end else
     if DansExo(V_PGI.EnCours,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.EnCours ; end else
         if DansExo(V_PGI.Suivant,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Suivant ; end;
  With Cum.FormatPrint Do
       begin
       PrSepMontant:=True ; PrSepCompte:=True ;
       PutTabCol(1, TabColl[1],-1, True) ;
       PutTabCol(2, TabColl[2], 2, True) ;
       PutTabCol(3, TabColl[3], 3, True) ;
       PutTabCol(4, TabColl[4], 5, True) ;
       PutTabCol(5, TabColl[5], 7, False) ;
       PutTabCol(6, TabColl[6], 9, False) ;
       end;
  end;
end;
*)

(*======================================================================*)
(*
procedure InitCritGL (Var CritEdt : TcritEdt ) ;
begin
With CritEdt Do
  begin
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ;
  DeviseSelect:='' ; DeviseAffichee:=V_PGI.DevisePivot ;
  Valide:='g' ; AfficheSymbole:=FALSE ;  MonoExo:=TRUE ;

  Gl.AnalPur:='g' ; Gl.TypCpt:=1 ;
  if V_PGI.EnCours.Deb=Date1 then
     begin
     Gl.Date11:=Date1 ; Gl.Date21:=Gl.Date11 ;
     end else
     begin
     Date1:=V_PGI.EnCours.Deb ; Gl.Date11:=DateDeb-1 ; Gl.Date21:=Gl.Date11+1 ;
     end;
  if DansExo(V_PGI.Precedent,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Precedent ; end else
     if DansExo(V_PGI.EnCours,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.EnCours ; end else
        if DansExo(V_PGI.Suivant,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Suivant ; end;
  With Gl.FormatPrint Do
    begin
    PrSepMontant:=True ;
    PrSepCompte[1]:=True ;
    PrSepCompte[2]:=True ;
    end;
  end;
end;
*)

(*======================================================================*)
(*
procedure InitCritBal (Var CritEdt : TcritEdt ) ;
begin
With CritEdt Do
  begin
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ;
  DeviseSelect:='' ; DeviseAffichee:=V_PGI.DevisePivot ;
  Valide:='g' ; AfficheSymbole:=FALSE ;  MonoExo:=TRUE ;

  Bal.AnalPur:='g' ; Bal.TypCpt:=1 ;
  if V_PGI.EnCours.Deb=Date1 then
     begin
     Bal.Date11:=Date1 ; Bal.Date21:=Bal.Date11 ;
     end else
     begin
     Date1:=V_PGI.EnCours.Deb ; Bal.Date11:=DateDeb-1 ; Bal.Date21:=Bal.Date11+1 ;
     end;
  if DansExo(V_PGI.Precedent,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Precedent ; end else
     if DansExo(V_PGI.EnCours,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.EnCours ; end else
        if DansExo(V_PGI.Suivant,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Suivant ; end;
  With Bal.FormatPrint Do
    begin
    PrSepMontant:=True ;
    PrSepCompte:=True ;
    PrSepRupt:=True ;
    PutTabCol(1, TabColl[1], -1, True) ;
    PutTabCol(2,TabColl[2], 2, True) ;
    PutTabCol(3,TabColl[3], 3, True) ;
    PutTabCol(4,TabColl[4], 5, True) ;
    PutTabCol(5,TabColl[5], 7, True) ;
    PutTabCol(6,TabColl[6], 9, True) ;
    end;
  end;
end;
*)

(*======================================================================*)
(*
procedure InitCritBrouil (Var CritEdt : TcritEdt ) ;
begin
With critEdt Do
  begin
  end;
end;
*)

(*======================================================================*)
(*
procedure InitCritGLVen (Var critEdt : TcritEdt ) ;
begin
With critEdt Do
  begin
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ; DeviseSelect:='' ;
  DeviseAffichee:=V_PGI.DevisePivot ;  AfficheSymbole:=FALSE;

  GlV.Ecart:=30 ;
  GlV.Sens:=2 ;
  With GlV.FormatPrint Do
    begin
    PrSepMontant:=True ;
    PrSepCompte[1]:=True ;
    Report.OkAff:=True ;
    end;
  end;
end;
*)

(*======================================================================*)
(*
procedure InitCritJustif (Var CritEdt : TcritEdt) ;
begin
With CritEdt Do
  begin
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ; DeviseSelect:='' ;
  DeviseAffichee:=V_PGI.DevisePivot ; AfficheSymbole:=False ;

  JU.Sens:=2 ;
  With JU.FormatPrint Do
       begin
       PrSepMontant:=True ;
       PrSepCompte[1]:=True ;
       PrSepCompte[2]:=True ;
       PrSepCompte[3]:=True ;
       PrSepCompte[4]:=True ;
       Report.OkAff:=False ;
       end;
  end;
end;
*)

(*======================================================================*)
(*
procedure InitCritEche (Var CritEdt : TcritEdt ) ;
begin
With CritEdt Do
  begin
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ; DeviseSelect:='' ;
  DeviseAffichee:=V_PGI.DevisePivot ;  MonoExo:=True ; AfficheSymbole:=False ;

  Ech.ModePaie:='' ;
  // Echeancier:=FEcheancier.ItemIndex ;
  Ech.Sens:=2 ;
  if DansExo(V_PGI.Precedent,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Precedent ; end else
     if DansExo(V_PGI.EnCours,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.EnCours ; end else
        if DansExo(V_PGI.Suivant,Date1,Date2) then begin MonoExo:=TRUE ; Exo:=V_PGI.Suivant ; end;
  With Ech.FormatPrint Do
       begin
       PrSepMontant:=True  ;
       PrSepCompte[1]:=True ;
       PrSepCompte[2]:=False ;
       Report.OkAff:=False ;
       end;
  end;
end;
*)

{================================================================}
{$IFNDEF EAGLCLIENT}
procedure ChgMaskQrCalc (C : TQRDBCalc ; Decim : Integer ; AffSymb : boolean ; Symbole : String ; IsSolde : Boolean) ;
Var j      : integer ;
begin
if Decim=0 then C.PrintMask:='#,##0' else C.PrintMask:='#,##0.' ;
for j:=1 to Decim do C.PrintMask:=C.PrintMask+'0' ;
if AffSymb then
   begin
   if IsSolde
      then begin C.PrintMask:=C.PrintMask+' '+Symbole+';-'+C.PrintMask+' '+Symbole+';'+C.PrintMask+' '+Symbole ; end
      else begin C.PrintMask:=C.PrintMask+' '+Symbole+';-'+C.PrintMask+' '+Symbole+';''''' ; end;
   end else
   begin
   if IsSolde
      then begin C.PrintMask:=C.PrintMask+' '+';-'+C.PrintMask+' '+';'+C.PrintMask ; end
      else begin C.PrintMask:=C.PrintMask+' '+';-'+C.PrintMask+' '+';''''' ; end;
   end;
end;

{================================================================}
procedure CASEACOCHER(FBox : TCheckBox ; RBox : TQRLabel);
begin
RBox.Font.Name:='WingDings' ; RBox.Font.Size:=10 ;
if FBox.State=cbUnChecked then begin RBox.Caption:='o' ; RBox.Font.Color:=clBlack ; end;
if FBox.State=cbChecked then begin RBox.Caption:='þ' ; RBox.Font.Color:=clBlack ; end;
if FBox.State=cbGrayed then begin RBox.Caption:='n' ; RBox.Font.Color:=clGray ; end;
end;

{================================================================}
procedure DATEBANDEAUBALANCE(Avant,Selection,Apres,Solde : TQRLABEL ; CritEdt : TCritEdt ; St0,St1,St2 : String) ;
Var SD0,SD1,SD2 : String ;
begin
SD0:=DateToStr(CritEdt.DateDeb-1) ;
SD1:=DateToStr(CritEdt.DateDeb) ;
SD2:=DateToStr(CritEdt.DateFin) ;
  // GC - 20/12/2001
  if CritEdt.AvecComparatif then
  begin
    Avant.Caption := St2 + ' ' + Sd2;
    if CritEdt.CompareBalSit then Selection.Caption := St0 else
    Selection.Caption := St2 + ' ' + DateToStr( CritEdt.ExoComparatif.Fin );

  end
  else
  begin
    Avant.Caption := St0 + ' ' + Sd0;
    Selection.Caption := St1 + ' ' + Sd1 + ' ' + St0 + ' ' + Sd2;
  end;
  // GC - 20/12/2001 - FIN
Apres.Caption:=St0+' '+Sd2;
 Solde.Caption:=St2+' '+Sd2;
end;

{================================================================}
procedure DATECUMULAUGL(CumulAu : TQRLABEL ; CritEdt : TCritEdt ; St0 : String) ;
Var SD0 : String ;
begin
SD0:=DateToStr(CritEdt.DateDeb-1) ;
CumulAu.Caption:=St0+' '+SD0 ;
end;
{$ENDIF}

{================================================================}
{ Gestion de Total Bilan, Charge et Produit  }
procedure ChargeBilan(Var L : TStringList) ;
Var Cpt : TCPTBCP ; i, j : integer ;
begin
L:=TStringList.Create ; L.Clear ;
Cpt:=TCPTBCP.Create ;
For i:=1 to high(VH^.FBIL) do begin Cpt.deb[i]:=VH^.FBIL[i].Deb ; Cpt.Fin[i]:=VH^.FBIL[i].Fin ; end;
Cpt.QuelCpt:='B' ; for j:=0 to High(Cpt.Total) do Cpt.Total[j]:=0; L.AddObject('',Cpt) ;

Cpt:=TCPTBCP.Create ;
For i:=1 to high(VH^.FCHA) do begin Cpt.deb[i]:=VH^.FCHA[i].Deb ; Cpt.Fin[i]:=VH^.FCHA[i].Fin ; end;
Cpt.QuelCpt:='C' ; for j:=0 to High(Cpt.Total) do Cpt.Total[j]:=0; L.AddObject('',Cpt) ;

Cpt:=TCPTBCP.Create ;
For i:=1 to high(VH^.FPRO) do begin Cpt.deb[i]:=VH^.FPRO[i].Deb ; Cpt.Fin[i]:=VH^.FPRO[i].Fin ; end;
Cpt.QuelCpt:='P' ; for j:=0 to High(Cpt.Total) do Cpt.Total[j]:=0; L.AddObject('',Cpt) ;

Cpt:=TCPTBCP.Create ;
For i:=1 to high(VH^.FExt) do begin Cpt.deb[i]:=VH^.FExt[i].Deb ; Cpt.Fin[i]:=VH^.FExt[i].Fin ; end;
Cpt.QuelCpt:='E' ; for j:=0 to High(Cpt.Total) do Cpt.Total[j]:=0; L.AddObject('',Cpt) ;
end;

{================================================================}
procedure AddBilan (L : TStringList ; LeCompte : String ; Montants : Array of double) ;
Var Cpt : TCPTBCP ; i, j, K : integer ;
begin
for i:=0 to L.Count-1 do
    begin
    Cpt:=TCPTBCP(L.Objects[i]) ;
    For J:=1 to high(Cpt.Deb) do
        begin
        if (Cpt.Deb[j]='')and(Cpt.Fin[j]='') then Break ;
        if (LeCompte>=Cpt.Deb[j])and(LeCompte<=Cpt.Fin[j]) then
           begin
           if Cpt.QuelCpt='B' then begin for K:=0 to High(Montants) do Cpt.Total[K]:=Cpt.Total[K]+Montants[K] ; end else
           if Cpt.QuelCpt='C' then begin for K:=0 to High(Montants) do Cpt.Total[k]:=Cpt.Total[K]+Montants[K] ; end else
           if Cpt.QuelCpt='P' then begin for K:=0 to High(Montants) do Cpt.Total[k]:=Cpt.Total[K]+Montants[K] ; end else
           if Cpt.QuelCpt='E' then begin for K:=0 to High(Montants) do Cpt.Total[k]:=Cpt.Total[K]+Montants[K] ; end;
           end;
        end;
    end;
end;

{================================================================}
procedure PrintBilan (L : TStringList ; Var Sommes : Array of double ; Qui : Char) ;
Var Cpt : TCPTBCP ; i, j : integer ;
begin
for j:=0 to High(Sommes) do Sommes[j]:=0;

for i:=0 to L.Count-1 do
    begin
    Cpt:=TCPTBCP(L.Objects[i]) ;
    if Cpt.QuelCpt=Qui then
       begin
       for j:=0 to High(Sommes) do Sommes[j]:=Cpt.Total[j];
       end;
    end;
end;

{================================================================}
procedure VideBilan(Var L : TStringList) ;
Var Cpt : TCPTBCP ; i : integer ;
begin
For i:=0 to L.Count-1 do
   begin
   Cpt:=TCPTBCP(L.Objects[i]) ;
   Cpt.Free ;
   end;
if L<>Nil then begin L.Clear ; L.Free ; end;
end;

(*======================================================================*)
function QuelleTypeRupt( i :Integer ; SAnsRupt,AvecRupt,SurRupt : Boolean ) : TRuptEtat ;
begin
Result:=Sans ;
if AvecRupt then Result:=Avec else if SurRupt And (i=0) then Result:=Sur ;
end;

{================================================================}
procedure DExoToDates ( Exo : String3 ; Var D1,D2 : TDateTime ) ;
Var Q     : TQuery;
begin
if EXO='' then Exit ;
D1:=Date ; D2:=Date ;
if EXO=VH^.Precedent.Code then begin D1:=VH^.Precedent.Deb ; D2:=VH^.Precedent.Fin ; end else
if EXO=VH^.EnCours.Code then begin D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ; end else
if EXO=VH^.Suivant.Code then begin D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; end else
   begin
   Q:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exo+'"' ,TRUE) ;
   if Not Q.EOF then
      begin
      D1:=Q.FindField('EX_DATEDEBUT').asDateTime ; D2:=Q.FindField('EX_DATEFIN').asDateTime ;
      end;
   Ferme(Q) ;
   end;
end;

{================================================================}
procedure VoirQuelAN(ValueLTypCpt : String ; CodeExo : String ; FD : TMaskEdit ; FAvecAN : TCheckBox) ;
Var DD1,DD2 : TDateTime ;
begin
if Not IsValidDate(FD.Text) then Exit ; // Rony 14/05/97
if ValueLTypCpt<>'PER' then begin FAvecAN.Visible:=FALSE ; FAvecAN.Checked:=TRUE ; Exit ; end;
DExoToDates(CodeExo,DD1,DD2) ;
if DD1<>StrToDate(FD.text) then FAvecAN.Visible:=TRUE else
   begin
   FAvecAN.Visible:=FALSE ; FAvecAN.Checked:=TRUE ;
   end;
end;

(*======================================================================*)
function ValiQuali(Valide, Qualif : String) : String ;
begin
if Valide='X' then result:='V' else
if Qualif='N' then result:='' else result:=Qualif ;
end;

(*======================================================================*)
function ValideTva(Valide : String) : String ;
begin
Result:='' ;
if Valide='X' then result:='P' else if Valide='#' then result:='V' ;
end;

{================================================================}
{$IFNDEF EAGLCLIENT}
procedure RedimLargeur(FF : TForm ; Qrp : TQuickReport ; Couleur : Boolean ; TTitre : TQRSysData);
var Ratio : double ;
    i,j   : integer ;
    C     : TQRCustomControl ;
begin
QRP.PrepareForPrinter ;

Ratio:=QRP.QRPrinter.PageWidth/TTitre.Width ; if Ratio=0 then Exit ;
For i:=0 to FF.ComponentCount-1 do
   begin
   if FF.Components[i] is TQRBand then
      begin
      For j:=0 to TQRBand(FF.Components[i]).ControlCount-1 do
          begin
          C:=TQRCustomControl(TQRBand(FF.Components[i]).Controls[j]) ;
          C.Left:=round(C.Left*Ratio) ; C.Width:=round(C.Width*Ratio) ;
          end;
      if ((TQRBand(FF.Components[i]).Name='BDCol') and Couleur) then ReajusteLabels(TQRBand(FF.Components[i])) ;
      end;
   end;
end;

{================================================================}
procedure ReajusteLabels(BDCol : TQRBand);
var T,i      : integer ;
    C1,C2    : TQRLabel ;
begin
C1:=Nil ; C2:=Nil ;
For T:=1 to MaxColListe-1 do
   begin
   For i:=0 to BDCol.ControlCount-1 do
      begin
      if BDCol.Controls[i] is TQRLabel then
         begin
         if TQRLabel(BDCol.Controls[i]).Tag=T then C1:=TQRLabel(BDCol.Controls[i]) else
         if TQRLabel(BDCol.Controls[i]).Tag=T+1 then C2:=TQRLabel(BDCol.Controls[i]) ;
         end;
      end;
   if C1<>C2 then C1.Width:=(C2.Left-C1.Left)-1 ;
   end;
end;
{$ENDIF}

(*======================================================================*)
function OrderLibre(LibreOrder : String ; SansVirgule : Boolean=FALSE) : String ;
Var St, StCode, StOrder : String ;
begin
St:=LibreOrder ; StCode:='' ; StOrder:='' ;
While St<>'' do
      begin
      StCode:=ReadTokenSt(St) ;
      if (StCode='') then Continue ;
      if Stcode[1]='B' then StOrder:=StOrder+Stcode[1]+'G_TABLE'+Stcode[3]+',' else
      if Stcode[1]='D' then StOrder:=StOrder+'BS_TABLE'+Stcode[3]+',' else
      StOrder:=StOrder+Stcode[1]+'_TABLE'+Stcode[3]+',' ;
      end;
if SansVirgule And (StOrder<>'') then Delete(StOrder,Length(StOrder),1) ;
Result:=StOrder ;
end;

(*======================================================================*)
function WhereLibre(LesCodes1, LesCodes2 : String ; LeFb : TFichierBase ; OnlyCpteAssocie : Boolean) : string ;
Var SqlSt : String ;
    Pref,C1,C2, LeTypeCpt, NatTyp : String ;
    i : Byte ;
begin
Result:='' ;
Case LeFb of
    fbGene               :  begin Pref:='G'  ; NatTyp:='G0' ; end;
    fbAux                :  begin Pref:='T'  ; NatTyp:='T0' ; end;
    fbImmo               :  begin Pref:='I'  ; NatTyp:='I0' ; end;
    fbBudgen             :  begin Pref:='BG' ; NatTyp:='B0' ; end;
    fbBudSec1..fbBudSec5 :  begin Pref:='BS' ; NatTyp:='D0' ; end;
    fbAxe1..fbAxe5       :  begin Pref:='S'  ; NatTyp:='S0' ; end;
    end;
i:=0 ; SQlSt:='' ;
//Exit ;
if (LesCodes1='') Or (LesCodes1='') then Exit ;
if Not OnlyCpteAssocie then Exit ;
While (LesCodes1<>'') do  //LesCodes1 == LesCodes2
      begin
      C1:=ReadTokenSt(LesCodes1) ; C2:=ReadTokenSt(LesCodes2) ;
      if C1='' then begin inc(i) ; continue ; end; //C1 == C2
      LeTypeCpt:=NatTyp+intToStr(i) ;
      (* GP le 22/12/97
      if ((C1[1]<>'#') AND (C1[1]<>'*') AND (C1[1]<>'-')) then SqlSt:=SqlSt+' and '+Pref+'_TABLE'+intToStr(i)+'>="'+C1+'" ' ;
      if ((C2[1]<>'#') AND (C2[1]<>'*') AND (C2[1]<>'-')) then SqlSt:=SqlSt+' and '+Pref+'_TABLE'+intToStr(i)+'<="'+C2+'" ' ;
      *)
      if (C1[1]<>'#') And (C2[1]<>'#') And (C1[1]<>'-') And (C2[1]<>'-') And (C1<>'') And (C2<>'') then
         begin
         if (C1[1]<>'*') And (C2[1]<>'*') then
            begin
            SqlSt:=SqlSt+' and '+Pref+'_TABLE'+intToStr(i)+'>="'+C1+'" ' ;
            SqlSt:=SqlSt+' and '+Pref+'_TABLE'+intToStr(i)+'<="'+C2+'" ' ;
            end else SqlSt:=SqlSt+' and '+Pref+'_TABLE'+intToStr(i)+'<>"" AND NOT ('+Pref+'_TABLE'+intToStr(i)+' IS NULL) ' ;
         end;
      inc(i) ;
      end;
//if SQLSt<>'' then SQLSt:='('+SQLSt+')' ;
Result:=SQLSt ;
end;

{================================================================}
procedure Reevaluation(Var D,C : Double ; Quelle : hString ; CoefRev : double ) ;
Var Coef : Double ;
    EDeb, ECre : Double ;
begin
//Coef:=1.0 ;
  EDeb:=D ; ECre:=C ;
  {JP 16/07/07 : connexe à la FQ 16050}
  if Length(Quelle) > 0 then
    Case Quelle[1] Of
      'K' : begin Coef:=1000.0 ; D:=(EDeb/Coef) ; C:=(ECre/Coef) ; end;
      'M' : begin Coef:=1000000.0 ; D:=(EDeb/Coef) ; C:=(ECre/Coef) ; end;
      end;
  if Arrondi(CoefRev,5)<>0 then begin D:=D*CoefRev ; C:=C*CoefRev ; end
end;

(*======================================================================*)
{$IFNDEF EAGLCLIENT}
function AfficheNewMontant( Var CritEdt : TCritEdt ; TB : TabDC ; MontantLabel : TQRLabel) : Boolean ;
Var D,C : Double ;
    Resol : hChar ;
    SoldeFormate : String3 ;
begin
D:=TB.TotDebit ; C:=TB.TotCredit ;
Result:=(Arrondi(D-C,4)=0) ;
Case CritEdt.NatureEtat Of
  neGLV : begin Resol:=CritEdt.GlV.Resol ; SoldeFormate:=CritEdt.GLV.SoldeFormate ; end;
  else Exit ;
  end;
Reevaluation(D,C,Resol,0) ;
MontantLabel.Caption:=PrintSoldeFormate(D,C, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
end;

{================================================================}
procedure AffecteLigne(FF : TForm ; Ent : TQRBAND ;CFP : TFormatPrintEDT) ;
Var T          : TComponent ;
    i,j        : Integer ;
    Gauche     : Integer ;
    LEnt, L    : TQRShape ;
begin
For j:=0 to FF.ComponentCount-1 do
    begin
    T:=FF.Components[J] ;
    if (T is TQRBand) and (TQRBand(T).Name<>Ent.Name) And
       (Uppercase(TQRBand(T).Name)<>'BLIBREHAUT') And (UpperCase(TQRBand(T).Name)<>'BRAPPELLIBREHAUT')then
       begin
       if (TQRBand(T).ControlCount)>CFP.ColMax then
       For i:=CFP.ColMin to CFP.ColMax do
           begin
           if (TQRBand(T).Controls[i].Tag>0) (*and TQRShape(TQRBand(T).Controls[i]).SynTitreCol*) then
              begin
              Gauche:=Ent.Controls[i].Left-1 ;
              L:=TQRShape.Create(ff) ;  L.Parent:=TQRBand(T) ;
              L.Shape:=qrsVertLine ; L.Name:='LVer'+InttoStr(J)+InttoStr(i) ;
              L.Tag:=i ;
              L.Top:=0 ; L.Height:=TQRBand(T).Height-2 ; L.Left:=Gauche ;
              end;
           end;

       end;
    end;
For i:=CFP.ColMin to CFP.ColMax do
    begin
    Gauche:=Ent.Controls[i].Left-1 ;
    {Entete page}
    LEnt:=TQRShape.Create(FF) ; LEnt.Name:='LEnt'+InttoStr(i) ;
    LEnt.Parent:=Ent ; LEnt.Shape:=qrsVertLine ; LEnt.Tag:=i ;
    LEnt.Top:=0 ; LEnt.Height:=Ent.Height-2 ; LEnt.Left:=Gauche ;
    end;
end;

{================================================================}
procedure FreeLesLignes(FF : TForm) ;
Var T   : TComponent ;
    I   : Integer ;
begin
I:=0;
while I<=FF.ComponentCount-1 do
    begin
    T:=FF.Components[i] ;
    if (T is TQRShape) then
       begin
       if (TQRShape(T).Tag>0) then TQRShape(T).Free else Inc(i) ;
       end else Inc(i) ;
    end;

end;
{$ENDIF}

(*======================================================================*)
function ExistBud(LeFb : TFichierBase ; Valeur, Jal, Axe : String ; First : Boolean ) : String ;
Var Pref, Pref2, Cpt, Cpt2, CXCpt, CXCpt2, st, Alias, Alias2 : String ;
begin
Case LeFb of
 fbBudgen             : begin Pref:='BG_' ; Cpt:='BUDGENE' ; CXCpt:='COMPTE' ; Pref2:='BS_' ; Cpt2:='BUDSECT' ; CXCpt2:='SECTION' ; Alias:='G.' ; Alias2:='S.' ; end;
 fbBudSec1..fbBudSec5 : begin Pref:='BS_' ; Cpt:='BUDSECT' ; CXCpt:='SECTION' ; Pref2:='BG_' ; Cpt2:='BUDGENE' ; CXCpt2:='COMPTE' ; Alias:='S.' ; Alias2:='G.' ;end;
  end;
St:='' ;
//Exit ;
if Valeur='CRO' then
   begin
   if First then St:='Exists (select CX_'+CXCpt+' from CROISCPT where CX_TYPE="BUD" AND CX_JAL="'+Jal+'" AND CX_'+CXCpt+'='+Alias+Pref+Cpt+') ' else
                 St:='Exists (select CX_'+CXCpt+', CX_'+CXCpt2+' from CROISCPT where CX_TYPE="BUD" AND CX_JAL="'+Jal+'" AND CX_'+CXCpt2+'=:'+Pref2+Cpt2+' AND CX_'+CXCpt+'='+Alias+Pref+Cpt+' ) ' ;
   end else
if Valeur='MVT' then
   begin
   if First then St:='Exists (select BE_'+Cpt+' from BUDECR where BE_'+Cpt+'='+Alias+Pref+Cpt+' and BE_BUDJAL="'+Jal+'" and BE_AXE="'+Axe+'" ) ' else
                 St:='Exists (select BE_'+Cpt+', BE_'+Cpt2+' from BUDECR where BE_'+Cpt2+'=:'+Pref2+Cpt2+'  AND BE_'+Cpt+'='+Alias+Pref+Cpt+' and BE_BUDJAL="'+Jal+'" and BE_AXE="'+Axe+'" ) ' ;
   end;
if Valeur='TOE' then
   begin
   if First then
      begin
      St:='((Exists (select BE_'+Cpt+' from BUDECR where BE_'+Cpt+'='+Alias+Pref+Cpt+' and BE_BUDJAL="'+Jal+'" and BE_AXE="'+Axe+'")) ' ;
      St:=St+' And (Exists (select CX_'+CXCpt+' from CROISCPT where CX_TYPE="BUD" AND CX_JAL="'+Jal+'" AND CX_'+CXCpt+'='+Alias+Pref+Cpt+'))) ' ;
      end else
      begin
      St:='((Exists (select BE_'+Cpt+', BE_'+Cpt2+' from BUDECR where BE_'+Cpt2+'=:'+Pref2+Cpt2+' and BE_'+Cpt+'='+Alias+Pref+Cpt+' and  BE_BUDJAL="'+Jal+'" and BE_AXE="'+Axe+'")) ' ;
      St:=St+' And (Exists (select CX_'+CXCpt+', CX_'+CXCpt2+' from CROISCPT where CX_TYPE="BUD" AND CX_JAL="'+Jal+'" AND CX_'+CXCpt2+'=:'+Pref2+Cpt2+' and CX_'+CXCpt+'='+Alias+Pref+Cpt+'  ))) ' ;
      end;
   end else
if Valeur='TOO' then
   begin
   if First then
      begin
      St:='((Exists (select BE_'+Cpt+' from BUDECR where BE_'+Cpt+'='+Alias+Pref+Cpt+' and BE_BUDJAL="'+Jal+'" and BE_AXE="'+Axe+'")) ' ;
      St:=St+' Or (Exists (select CX_'+CXCpt+' from CROISCPT where CX_TYPE="BUD" AND CX_JAL="'+Jal+'" AND CX_'+CXCpt+'='+Alias+Pref+Cpt+'))) ' ;
      end else
      begin
      St:='((Exists (select BE_'+Cpt+', BE_'+Cpt2+' from BUDECR where BE_'+Cpt2+'=:'+Pref2+Cpt2+' and BE_'+Cpt+'='+Alias+Pref+Cpt+' and BE_BUDJAL="'+Jal+'" and BE_AXE="'+Axe+'")) ' ;
      St:=St+' Or (Exists (select CX_'+CXCpt+', CX_'+CXCpt2+' from CROISCPT where CX_TYPE="BUD" AND CX_JAL="'+Jal+'" AND CX_'+CXCpt2+'=:'+Pref2+Cpt2+' AND CX_'+CXCpt+'='+Alias+Pref+Cpt+' ))) ' ;
      end;
   end;
Result:=St ;
end;

(*======================================================================*)
function ExistBudOle(LeFb : TFichierBase ; Valeur, Jal, Axe : String ; First : Boolean ; AL1, AL2 : String ) : String ;
Var Pref, Pref2, Cpt, Cpt2, CXCpt, CXCpt2, st, Alias, Alias2 : String ;
begin
Case LeFb of
 fbBudgen             : begin Pref:='BG_' ; Cpt:='BUDGENE' ; CXCpt:='COMPTE' ; Pref2:='BS_' ; Cpt2:='BUDSECT' ; CXCpt2:='SECTION' ; Alias:=AL1 ; Alias2:=AL2 ; end;
 fbBudSec1..fbBudSec5 : begin Pref:='BS_' ; Cpt:='BUDSECT' ; CXCpt:='SECTION' ; Pref2:='BG_' ; Cpt2:='BUDGENE' ; CXCpt2:='COMPTE' ; Alias:=AL1 ; Alias2:=AL2 ;end;
  end;
St:='' ;
//Exit ;
if Valeur='CRO' then
   begin
   if First then St:='Exists (select CX_'+CXCpt+' from CROISCPT where CX_TYPE="BUD" AND CX_JAL="'+Jal+'" AND CX_'+CXCpt+'='+Alias+Pref+Cpt+') ' else
                 St:='Exists (select CX_'+CXCpt+', CX_'+CXCpt2+' from CROISCPT where CX_TYPE="BUD" AND CX_JAL="'+Jal+'" AND CX_'+CXCpt2+'=:'+Pref2+Cpt2+' AND CX_'+CXCpt+'='+Alias+Pref+Cpt+' ) ' ;
   end else
if Valeur='MVT' then
   begin
   if First then St:='Exists (select BE_'+Cpt+' from BUDECR where BE_'+Cpt+'='+Alias+Pref+Cpt+' and BE_BUDJAL="'+Jal+'" and BE_AXE="'+Axe+'" ) ' else
                 St:='Exists (select BE_'+Cpt+', BE_'+Cpt2+' from BUDECR where BE_'+Cpt2+'=:'+Pref2+Cpt2+'  AND BE_'+Cpt+'='+Alias+Pref+Cpt+' and BE_BUDJAL="'+Jal+'" and BE_AXE="'+Axe+'" ) ' ;
   end;
if Valeur='TOE' then
   begin
   if First then
      begin
      St:='((Exists (select BE_'+Cpt+' from BUDECR where BE_'+Cpt+'='+Alias+Pref+Cpt+' and BE_BUDJAL="'+Jal+'" and BE_AXE="'+Axe+'")) ' ;
      St:=St+' And (Exists (select CX_'+CXCpt+' from CROISCPT where CX_TYPE="BUD" AND CX_JAL="'+Jal+'" AND CX_'+CXCpt+'='+Alias+Pref+Cpt+'))) ' ;
      end else
      begin
      St:='((Exists (select BE_'+Cpt+', BE_'+Cpt2+' from BUDECR where BE_'+Cpt2+'=:'+Pref2+Cpt2+' and BE_'+Cpt+'='+Alias+Pref+Cpt+' and  BE_BUDJAL="'+Jal+'" and BE_AXE="'+Axe+'")) ' ;
      St:=St+' And (Exists (select CX_'+CXCpt+', CX_'+CXCpt2+' from CROISCPT where CX_TYPE="BUD" AND CX_JAL="'+Jal+'" AND CX_'+CXCpt2+'=:'+Pref2+Cpt2+' and CX_'+CXCpt+'='+Alias+Pref+Cpt+'  ))) ' ;
      end;
   end else
if Valeur='TOO' then
   begin
   if First then
      begin
      St:='((Exists (select BE_'+Cpt+' from BUDECR where BE_'+Cpt+'='+Alias+Pref+Cpt+' and BE_BUDJAL="'+Jal+'" and BE_AXE="'+Axe+'")) ' ;
      St:=St+' Or (Exists (select CX_'+CXCpt+' from CROISCPT where CX_TYPE="BUD" AND CX_JAL="'+Jal+'" AND CX_'+CXCpt+'='+Alias+Pref+Cpt+'))) ' ;
      end else
      begin
      St:='((Exists (select BE_'+Cpt+', BE_'+Cpt2+' from BUDECR where BE_'+Cpt2+'=:'+Pref2+Cpt2+' and BE_'+Cpt+'='+Alias+Pref+Cpt+' and BE_BUDJAL="'+Jal+'" and BE_AXE="'+Axe+'")) ' ;
      St:=St+' Or (Exists (select CX_'+CXCpt+', CX_'+CXCpt2+' from CROISCPT where CX_TYPE="BUD" AND CX_JAL="'+Jal+'" AND CX_'+CXCpt2+'=:'+Pref2+Cpt2+' AND CX_'+CXCpt+'='+Alias+Pref+Cpt+' ))) ' ;
      end;
   end;
Result:=St ;
end;

{================================================================}
procedure InitResolution(FRESOLUTION: THValComboBox) ;
begin
FResolution.ItemIndex:=1 ;
end;


{================================================================}
{$IFNDEF EAGLCLIENT}
{$IFDEF EDTQR}
procedure AlimRuptSup(HauteurBandeRuptIni,QuelleRupt,NewWidth : Integer ; BRupt : TQRBAND ;
                      LibRuptInf : Array Of TRuptInf ; FF : TForm) ;
Var i : Integer ;
    TC : TComponent ;
    T : TQRLabel ;
begin
//Exit ;
if BRupt=NIL then Exit ;
if (HauteurBandeRuptIni=0) then Exit ;
BRupt.Height:=HauteurBandeRuptIni ;
(*
if V_PGI.SAV And V_PGI.VersionDev then
   begin
*)
//   BRupt.Height:=HauteurBandeRuptIni*(QuelleRupt+1) ;
   BRupt.Height:=HauteurBandeRuptIni+20*(QuelleRupt) ;
   For i:=1 To 10  Do
     begin
     TC:=FF.FindComponent('TCODRUPT'+IntToStr(i)) ;
     if TC<>NIL then
        begin
        T:=TQRLABEL(TC) ;
        if NewWidth>0 then T.Width:=NewWidth ; T.Visible:=TRUE ;
        if i<=QuelleRupt then T.Caption:=LibRuptInf[QuelleRupt-i].CodRupt+LibRuptInf[QuelleRupt-i].LibRupt
                         else begin T.Caption:='' ; T.Visible:=FALSE ; end;
        end;
     end;
(*
   end;
*)
end;

(*======================================================================*)
function AlimStRuptSup(QuelleRupt : Integer ; Cod,Lib : String ; LibRuptInf : Array Of TRuptInf) : String ;
Var i : Integer ;
    St : String ;
begin
//' ('+StCode+' '+Lib1+')'
St:='' ;
if (Cod<>'') Or (Lib<>'') then St:=Cod+' '+Lib ;
For i:=1 To 10  Do
  begin
  if i<=QuelleRupt then St:=St+'/'+LibRuptInf[QuelleRupt-i].CodRupt+' '+LibRuptInf[QuelleRupt-i].LibRupt ;
  end;
if St<>'' then St:=' ('+St+')' ;
Result:=St ;
end;
{$ENDIF}
{$ENDIF}

{================================================================}
procedure StPourMontantsBudgetEcart(M : TabTot12 ; Var CritEdt : TCritEdt ; DebitPos : Boolean ; var AffM : array of String) ;
Var EcartPer, EcartCum, Pourcentage,Dper,CPer,DCum,CCum : Double ;
begin
DPer:=M[1].TotDebit-M[0].TotDebit ; CPer:= M[1].TotCredit-M[0].TotCredit ;
DCum:=M[5].TotDebit-M[4].TotDebit ; CCum:= M[5].TotCredit-M[4].TotCredit ;
EcartPer:=Arrondi((M[0].TotDebit-M[0].TotCredit)-(M[1].TotDebit-M[1].TotCredit), CritEdt.Decimale) ;
EcartCum:=Arrondi((M[4].TotDebit-M[4].TotCredit)-(M[5].TotDebit-M[5].TotCredit), CritEdt.Decimale) ;
{Total compte en période}
AffM[0]:=PrintSoldeFormate(M[0].TotDebit, M[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
AffM[1]:=PrintSoldeFormate(M[1].TotDebit, M[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;

//AffM[2]:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(EcartPer), CritEDT.AfficheSymbole ) ;
AffM[2]:=PrintEcart(DPer,Cper,CritEDT.Decimale,DebitPos) ;
if (AffM[0]<>'') and (AffM[2]<>'') then
   begin
   Pourcentage:=EcartPer*100/(M[0].TotDebit-M[0].TotCredit) ;
   Pourcentage:=Abs(Pourcentage) ;
   if Pos('-',AffM[2])>0 then Pourcentage:=-1*Pourcentage ;
   end else Pourcentage:=0 ;
AffM[3]:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Pourcentage , CritEDT.AfficheSymbole ) ;
{Total compte en Cumulé}
AffM[4]:=PrintSoldeFormate(M[4].TotDebit, M[4].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
AffM[5]:=PrintSoldeFormate(M[5].TotDebit, M[5].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
//AffM[6]:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(EcartCum), CritEDT.AfficheSymbole ) ;
AffM[6]:=PrintEcart(DCum,CCum,CritEDT.Decimale,DebitPos ) ;
if (AffM[4]<>'') and (AffM[6]<>'') then
   begin
   Pourcentage:=EcartCum*100/(M[4].TotDebit-M[4].TotCredit) ;
   Pourcentage:=Abs(Pourcentage) ;
   if Pos('-',AffM[6])>0 then Pourcentage:=-1*Pourcentage ;
   end else Pourcentage:=0 ;
AffM[7]:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Pourcentage , CritEDT.AfficheSymbole ) ;
{Total compte en Annuel}
AffM[8]:=PrintSoldeFormate(M[8].TotDebit, M[8].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
{Total compte en Extrapolé}
AffM[9]:='' ;
end;

{================================================================}
procedure AlimTotEdtEcartBudget(Var LeTotal : TabTot12 ; Var SousTot : TMontTot ; Var CritEdt : TCritEdt) ;
begin
LeTotal[0].TotDebit:= Arrondi(LeTotal[0].TotDebit+(SousTot[0][0].totDebit+SousTot[1][0].totDebit), CritEdt.Decimale) ;
LeTotal[0].TotCredit:= Arrondi(LeTotal[0].TotCredit+(SousTot[0][0].totCredit+SousTot[1][0].totCredit), CritEdt.Decimale) ;
LeTotal[4].TotDebit:= Arrondi(LeTotal[4].TotDebit+(SousTot[0][1].totDebit+SousTot[1][1].totDebit), CritEdt.Decimale) ;
LeTotal[4].TotCredit:= Arrondi(LeTotal[4].TotCredit+(SousTot[0][1].totCredit+SousTot[1][1].totCredit), CritEdt.Decimale) ;
LeTotal[8].TotDebit:= Arrondi(LeTotal[8].TotDebit+(SousTot[0][2].totDebit+SousTot[1][2].totDebit), CritEdt.Decimale) ;
LeTotal[8].TotCredit:= Arrondi(LeTotal[8].TotCredit+(SousTot[0][2].totCredit+SousTot[1][2].totCredit), CritEdt.Decimale) ;
end;

{================================================================}
{$IFNDEF EAGLCLIENT}
procedure AfficheBudgetEn(RDevises : TQRLabel ; Var CritEdt : TCritEdt) ;
begin
Case CritEdt.Monnaie Of
  0 : RDevises.Caption:=VH^.LibDevisePivot ;
  1 : RDevises.Caption:='' ;
  end;
end;
{$ENDIF}

(*======================================================================*)
function InitDecimaleResol(Resol : hChar ; Decimale : Integer) : Integer ;
begin
Result:=Decimale ;
if Resol='C' then Result:=Decimale else
if Resol='F' then Result:=0 else
if Resol='K' then Result:=0 else
if Resol='M' then Result:=0 ;
end;

{================================================================}
procedure SwapSymbolePivotDevise(Var CritEdt : TCritEdt) ;
Var SauvSymbole : String ;
    SauvDecimale : Integer ;
begin
SauvSymbole:=CritEdt.Symbole ;
SauvDecimale:=CritEdt.Decimale ;
CritEdt.Symbole:=CritEdt.SymbolePivot ;
CritEdt.Decimale:=CritEdt.DecimalePivot ;
CritEdt.SymbolePivot:=SauvSymbole ;
CritEdt.DecimalePivot:=SauvDecimale ;
end;

(*======================================================================*)
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 08/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function SQLCollCliEncTVA(Alias : String) : String;
var
    S,Sc : string ;
    St : String ;
begin
    if (Alias <> '') then if (Alias[Length(Alias)] <> '.') then Alias := Alias + '.';

    S := VH^.CollCliEnc;
    if (S = '') then
    begin
        result := 'AND (' + Alias + 'E_GENERAL="W_W")';
        exit;
    end;

    St := 'AND (';
    while (S <> '') do
    begin
        Sc := ReadTokenSt(s);
        if (St <> 'AND (') then St := St + ' OR ';
        St := St + Alias + 'E_GENERAL LIKE "' + Sc + '%"';
    end;

    result := St + ')';
end;

(*======================================================================*)
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 08/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function SQLCollFouEncTVA(Alias : String) : String;
var
    S,Sc : string ;
    St : String ;
begin
    if (Alias <> '') then if (Alias[Length(Alias)] <> '.') then Alias := Alias + '.';

    S := VH^.CollFouEnc;
    if (S = '') then
    begin
        result := 'AND (' + Alias + 'E_GENERAL="W_W")';
        exit;
    end;

    St := 'AND (';
    while (S <> '') do
    begin
        Sc := ReadTokenSt(s);
        if (St <> 'AND (') then St := St + ' OR ';
        St := St + Alias + 'E_GENERAL LIKE "' + Sc + '%"';
    end;

    result := St + ')';
end;

(*======================================================================*)
{$IFNDEF EAGLCLIENT}
function SautPageRuptAFaire(Var CritEdt : tCritEdt ; QRB : TQRBand ; QuelleRupt : Integer) : Boolean ;
begin
Result:=FALSE ;
if CritEdt.SautPageRupt then
  Case CritEdt.Rupture Of
    rLibres : begin QRB.ForceNewPage:=TRUE ; Result:=TRUE ; end;
    rRuptures : if QuelleRupt=0 then begin QRB.ForceNewPage:=TRUE ; Result:=TRUE ; end;
    rCorresp : begin QRB.ForceNewPage:=TRUE ; Result:=TRUE ; end;
  end;
end;
{$ENDIF}

(*======================================================================*)
function TrouveLibTP(St : String ; OkLib : Boolean) : String ;
Var StLib : String ;
    QAux : TQuery ;
begin
StLib:='' ;
if St<>'' then
  begin
  if OkLib then QAux:=OpenSQL('SELECT T_ABREGE FROM TIERS WHERE T_AUXILIAIRE="'+St+'" ',TRUE)
           else QAux:=OpenSQL('SELECT T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="'+St+'" ',TRUE) ;
  if Not QAux.Eof then StLib:=QAux.Fields[0].AsString else StLib:=TraduireMemoire('Inexistant') ;
  Ferme(QAux) ;
  end;
Result:=StLib ;
end;

{================================================================}
{$IFNDEF EAGLCLIENT}
procedure LanceDoc(Tip,Nat,Modele : String ; TT : Tlist ; Load : TLoad_Etiq ; Apercu,PrintDialog : boolean) ;
Var TT1 : TList ;
    LL : tStringList ;
    i : Integer ;
begin
for i:=0 to TT.Count-1 do
  begin
  TT1:=TList.Create ; LL:=TStringList.Create ;
  TT1.Add(TStringList(TT[i])) ;
  (*
  LL.assign(TStringList(TT[i])) ; TT1.Add(LL) ;
  *)
  LanceDocument(Tip,Nat,MODELE,TT1,Load,Apercu,PrintDialog) ;
  LL.Free; TT1.Free ;
  end;
end;
{$ENDIF}


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 09/11/2006
Modifié le ... :   /  /    
Description .. : Alimente le LOG d'info des éditions légales.
Suite ........ : Déplacée dans Utiledt car utilisée dans toutes les unités 
Suite ........ : utilisées pour les éditions légales.
Mots clefs ... : 
*****************************************************************}
Procedure MajEditionLegal(LeType,Exo,D1,D2: String) ;
Var Q : TQuery ;
    Legale : Boolean ;
    VirtExo : TExoDate ;
    D1Us,D2Us,DjUs,NomUser,St,St1 : String ;
    SCd,SCc,SSd,SSc : String ;
    M,A,NbrMois : Word ;
    DFMois,LeD1,LeD2 : TDateTime ;
    i,Num : Integer ;
    Dest : String ; Cd,Cc,Sd,Sc : Extended;
BEGIN
{$IFDEF CCS3}
  Exit ;
{$ENDIF}
  Dest := '';

  D1Us:=UsDateTime(StrToDateTime(D1)) ; D2Us:=UsDateTime(StrToDateTime(D2)) ;
  DjUs:=UsDateTime(Date) ; NomUser:=V_PGI.User ;

  St := 'Select CO_ABREGE From COMMUN Where CO_TYPE="EDL" And CO_CODE="'+LeType+'"';
  Q :=  OpenSQL(St, True);
  Legale := (Q.Fields[0].AsString='X') ;
  Ferme(Q);

  if Legale then
    BEGIN
    BeginTrans ;
    Cd:=0 ; Cc:=0 ; Sd:=0 ; Sc:=0 ;
    SCd:=StrfPoint(Cd) ; SCc:=StrfPoint(Cc) ; SSd:=StrfPoint(Sd) ; SSc:=StrfPoint(Sc) ;

    DFMois:=DebutDeMois(StrToDate(D1)) ;
    if(DFMois<StrTodate(D1)) then DFMois:=PlusMois(DFMois,1) ; LeD1:=DFMois ;

    DFMois:=FinDeMois(StrToDate(D2)) ;
    if(DFMois>StrTodate(D2)) then DFMois:=PlusMois(DFMois,-1) ; LeD2:=DFMois ;

    VirtExo.Deb:=LeD1 ; VirtExo.Fin:=LeD2 ; NombrePerExo(VirtExo,M,A,NbrMois) ;
    for i:=1 to NbrMois do
      BEGIN
      D1Us:=UsDateTime(LeD1) ; D2Us:=UsDateTime(FinDeMois(LeD1)) ;

      St:='Select ED_NUMEROEDITION From EDTLEGAL Where ED_TYPEEDITION="'+LeType+'"'+
          ' AND ED_EXERCICE="'+Exo+'" AND ED_OBLIGATOIRE="X"' +
          ' AND ED_PERIODE="'+D1Us+'"';
      Q := OpenSQL(St, True);
      if Q.Eof then Num:=1 else Num:=Q.Fields[0].AsInteger+1 ;
      
      if Num=1 then
        St1:='INSERT INTO EDTLEGAL (ED_OBLIGATOIRE, ED_TYPEEDITION, ED_EXERCICE, ED_DATE1,'+
             'ED_DATE2, ED_DATEEDITION, ED_NUMEROEDITION, ED_PERIODE, ED_CUMULDEBIT, '+
             'ED_CUMULCREDIT, ED_CUMULSOLDED, ED_CUMULSOLDEC, '+
             'ED_UTILISATEUR, ED_DESTINATION)'+
             ' VALUES ("X", "'+LeType+'", "'+Exo+'", "'+D1Us+'", "'+D2Us+'", "'+DjUs+'", '+
                       IntToStr(Num)+', "'+D1Us+'", '+SCd+', '+SCc+', '+
                       SSd+', '+SSc+', "'+NomUser+'", "'+Dest+'")'
      else
        St1:='UPDATE EDTLEGAL SET ED_OBLIGATOIRE="X", ED_TYPEEDITION="'+LeType+'", '+
             'ED_EXERCICE="'+Exo+'", ED_DATE1="'+D1Us+'", ED_DATE2="'+D2Us+'", ED_DATEEDITION="'+DjUs+'", '+
             'ED_NUMEROEDITION='+IntToStr(Num)+', ED_PERIODE="'+D1Us+'", ED_CUMULDEBIT='+SCd+', '+
             'ED_CUMULCREDIT='+SCc+', ED_CUMULSOLDED='+SSd+', '+
             'ED_CUMULSOLDEC='+SSc+', ED_UTILISATEUR="'+NomUser+'", ED_DESTINATION="'+Dest+'" '+
             'Where ED_TYPEEDITION="'+LeType+'" AND ED_EXERCICE="'+Exo+'" '+
             'AND ED_OBLIGATOIRE="X" AND ED_PERIODE="'+D1Us+'"' ;
       Ferme(Q);
       ExecuteSQL(St1) ;

       LeD1:=PlusMois(LeD1,1) ;
      END ;
    CommitTrans ;
    END ;
  Ferme(Q) ;
END ;

end.
