{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 12/10/2001
Modifié le ... :   /  /
Description .. : Fonctions utiles à la paie..
Mots clefs ... : PAIE;OUTILS
*****************************************************************
PT1   : 12/10/2001 SB V562 Edition d'un alphanumérique sur les attestations
                           erronées
                           Ajout ds la const de l'alphabet  fiche bug n°342
PT2   : 28/01/2002 SB V571 Procedure de sélection alphanumérique du code salarié
PT3   : 29/03/2002 SB V571 Fn Visibilite utilisé dans des fiches autres que QRS1
PT4   : 10/12/2002 SB V591 Modifcation FormatCase suite changement PDF DUE
PT5   : 05/08/2003 SB V_42 FQ 10708 Déactivation de l'affectation alphabétique
                           sur les fiches de lancement d'état
PT6   : 05/04/2004 PH V_50 Prise en compte des états des provisions cp et Rtt
PT7   : 24/09/2004 VG V_50 Déplacement de RecupMinMaxTablette dans PGEditOutils2
}

unit PgEditOutils;

interface
uses  StdCtrls,
      Controls,
      HCtrls,
      forms,
      sysutils,
{$IFDEF EAGLCLIENT}
      eQRS1,
      UTOB,
{$ELSE}
      db,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      QRS1,
{$ENDIF}
      Ed_tools,
      PGEditOutils2 ;

Var RecupOrderBy : string;
{------------------------------------------------------------------------------
                                 DECLARATION}
//Affiche les zones Lieu de Travail en fonction du paramétrage Sociétès associès
procedure VisibiliteChamp (Ecran : TForm);
//S'execute sur l'evenement On change de toute fiche QRS1 référencant les Lieux de Travail
procedure BloqueChampRupture(Ecran:TForm);
procedure MiseEnPageRupture(Ecran : TForm);
procedure RecupChampRupture(Ecran : TForm);
function  AffectSqlEtat(Ecran:TForm)  : string;
function  RendTauxConvertion : string;
function  FormatCase(St, StType: string ; Esp : integer): string;
Function  RemplaceVirguleParPoint (St : String) : String;
procedure VisibiliteChampLibre(Ecran : TForm);
procedure BloqueChampLibre(Ecran:TForm);
function  PGRendPrefixeFromQrs1 (Ecran:TForm; Tip,Nat,Modele : String): string;
Procedure AffectCritereAlpha(Ecran : TForm ;TriAlpha : Boolean; Champ1,Champ2 : String);
function  RendRequeteSQLEtat(NomEtat,Champ,Critere,OrderBy : string) : string;

{------------------------------------------------------------------------------}

implementation

uses 
     pgOutils,
     EntPaie;

var PrefTable,Tip,Nat,Modele : string;


procedure VisibiliteChamp (Ecran : TForm);
var
TLieu,TA,TCodeStat   : THLabel ;      // V_42 Qualité suppression variable
ChampLieu,ChampLieu_ : THEdit;
num,nr : integer;
numero,Min,Max,libelle : string;
Rupture:TCheckBox;
begin
if Ecran is TFQRS1 then //PT3 PrefTable forcé à PPU si fiche <> QRS1
  Begin
  Tip:=TFQRS1(Ecran).TypeEtat;
  Nat:=TFQRS1(Ecran).NatureEtat;
  Modele:=TFQRS1(Ecran).CodeEtat;
  PrefTable := PGRendPrefixeFromQrs1 (Ecran,Tip,Nat,Modele);
  if nat='PSO' then PrefTable:='PHB';
// PT6 05/04/2004 PH V_50 Prise en compte des états des provisions cp et Rtt
  if (nat='PCG') and ((Modele='PRT') OR (Modele='PDP') OR (Modele='PRP')) then PrefTable:='PSA';
  end
else  PrefTable:='PPU';

TCodeStat:=THLabel(PGGetControl(('T'+PrefTable+'_CODESTAT'),Ecran));
For num := 1 to VH_Paie.PGNbreStatOrg do
  begin
  Numero:=InttoStr(num);
  if Num > VH_Paie.PGNbreStatOrg then exit;
  Libelle:='';
  if Num = 1 then libelle:=VH_Paie.PGLibelleOrgStat1;
  if Num = 2 then libelle :=VH_Paie.PGLibelleOrgStat2;
  if Num = 3 then libelle :=VH_Paie.PGLibelleOrgStat3;
  if Num = 4 then libelle :=VH_Paie.PGLibelleOrgStat4;
  if (Libelle<>'') then
    Begin
    Min:='';Max:='';
//    if Num = 1 then RecupMinMaxTablette('CC','','PAG',Min,Max);
//    if Num = 2 then RecupMinMaxTablette('CC','','PST',Min,Max);
//    if Num = 3 then RecupMinMaxTablette('CC','','PUN',Min,Max);
//    if Num = 4 then RecupMinMaxTablette('CC','','PSV',Min,Max);
    ChampLieu:=THEdit(PGGetControl((PrefTable+'_TRAVAILN'+Numero),Ecran));
    if ChampLieu <> NIL then begin ChampLieu.Visible :=TRUE; ChampLieu.text:=Min; end;
    ChampLieu_:=THEdit(PGGetControl((PrefTable+'_TRAVAILN'+Numero+'_'),Ecran));
    if ChampLieu_ <> NIL then begin ChampLieu_.Visible :=TRUE; ChampLieu_.text:=Max;end;
    Rupture:=TCheckBox(PGGetControl(('CN'+Numero),Ecran));
    if Rupture <> NIL then Rupture.Visible :=TRUE;
    TA:=THLabel(PGGetControl (('TA'+Numero),Ecran));
    if TA <> NIL then TA.visible:=True;
    TLieu:=THLabel(PGGetControl(('T'+PrefTable+'_TRAVAILN'+Numero),Ecran));
    if (TLieu <> NIL)   then
      begin
      TLieu.Visible :=TRUE;
      if Num = 1 then Begin TLieu.Caption :=VH_Paie.PGLibelleOrgStat1+' de'; end;
      if Num = 2 then Begin TLieu.Caption :=VH_Paie.PGLibelleOrgStat2+' de'; end;
      if Num = 3 then Begin TLieu.Caption :=VH_Paie.PGLibelleOrgStat3+' de'; end;
      if Num = 4 then Begin TLieu.Caption :=VH_Paie.PGLibelleOrgStat4+' de'; end;
      end;
    end;
  end;
  If  VH_Paie.PGNbreStatOrg<4 then
     For Nr:=VH_Paie.PGNbreStatOrg to 4 do
       Begin
       TLieu:=THLabel(PGGetControl(('T'+PrefTable+'_TRAVAILN'+InttoStr(Nr+1)),Ecran));
       if (TLieu <> NIL)   then TLieu.Caption :='';
       End;

  //Visibilité du code statistique
  if VH_Paie.PGLibCodeStat<>'' then
    begin
    ChampLieu:=THEdit(PGGetControl(PrefTable+'_CODESTAT',Ecran));
    ChampLieu_:=THEdit(PGGetControl(PrefTable+'_CODESTAT_',Ecran));
    Rupture:=TCheckBox(PGGetControl('CN5',Ecran));
    TA:=THLabel(PGGetControl (('TA5'),Ecran));
    If (TCodeStat<>nil) and (ChampLieu<>nil) and (ChampLieu_<>nil)  then
       begin
       Min:=''; Max:='';
      // RecupMinMaxTablette('CC','','PSQ',Min,Max);
       TCodeStat.Visible:=True;
       TCodeStat.caption := VH_Paie.PGLibCodeStat+' de';
       ChampLieu.Visible:=True;
       ChampLieu.text:=Min;
       ChampLieu_.Visible:=True;
       ChampLieu_.text:=Max;
       if TA <> NIL then TA.visible:=True;
       if (rupture<>nil) then rupture.Visible:=True;
       end;
    end;
    MiseEnPageRupture(Ecran);
end;


procedure BloqueChampRupture(Ecran:TForm);
var
TabLieuTravail : array[1..6] of TCheckBox;
Min,Max : string;
PosCheck,PosUnCheck,i : integer;
BEGIN
TabLieuTravail[1]:=TCheckBox(PGGetControl('CN1',Ecran));
TabLieuTravail[2]:=TCheckBox(PGGetControl('CN2',Ecran));
TabLieuTravail[3]:=TCheckBox(PGGetControl('CN3',Ecran));
TabLieuTravail[4]:=TCheckBox(PGGetControl('CN4',Ecran));
TabLieuTravail[5]:=TCheckBox(PGGetControl('CN5',Ecran));
TabLieuTravail[6]:=TCheckBox(PGGetControl('CETAB',Ecran));
PosUnCheck:=0;
PosCheck:=0;
if (TabLieuTravail[1]<>nil) and (TabLieuTravail[2]<>nil) then
if (TabLieuTravail[3]<>nil) and (TabLieuTravail[4]<>nil) and (TabLieuTravail[5]<>nil) then
  begin
  //Coche une rupture
  For i:=1 to 6 do
     if (TabLieuTravail[i].checked=True) then PosCheck:=i;
  if PosCheck<>0 then
    begin
     For i:=1 to 6 do  //rend unenable(false) les autres champs de rupture
        if i<>PosCheck then
         TabLieuTravail[i].enabled:=False;
    end;
     //Décoche une rupture ,  rend enable(True) les autres champs de rupture
  For i:=1 to 6 do
    if (TabLieuTravail[i].checked=False) and (TabLieuTravail[i].enabled=True) then
      PosUnCheck:=i;
  if PosUnCheck<>0 then
    begin
    For i:=1 to 6 do
      begin
      if i<>PosUnCheck then
        begin
        TabLieuTravail[i].enabled:=True;
        Min:=''; Max:='';
        end;
      end;
    end;
  end;
END;

procedure MiseEnPageRupture(Ecran : TForm);
var
  CodeStat,CodeStat_ : THEdit;
  TCodeStat : THLabel;
  CN: TCheckBox;
begin
CodeStat:=THEdit(PGGetControl(PrefTable+'_CODESTAT',Ecran));
CodeStat_:=THEdit(PGGetControl(PrefTable+'_CODESTAT_',Ecran));
Cn:= TCheckBox(PGGetControl('CN5',Ecran));
TcodeStat:=THLabel(PGGetControl('T'+PrefTable+'_CODESTAT',Ecran));
if (VH_Paie.PGNbreStatOrg<3) and (VH_Paie.PGLibCodeStat<>'') then
  begin
  if (CodeStat<>nil) and (CodeStat_<>nil) and (TCodeStat<>nil)  then
     begin CodeStat.top:=52; CodeStat_.top:=52;  TCodeStat.top:=56; end;
  if (Cn<>nil) then  Cn.top:=54;
  end;
end;

procedure RecupChampRupture(Ecran : TForm);
var
CEtab,CN1,CN2,CN3,CN4,CN5,CL1,CL2,CL3,CL4:TCheckBox;
Rupture,Champ1 : THEdit;
begin   
CN1:=TCheckBox(PGGetControl('CN1',Ecran));
CN2:=TCheckBox(PGGetControl('CN2',Ecran));
CN3:=TCheckBox(PGGetControl('CN3',Ecran));
CN4:=TCheckBox(PGGetControl('CN4',Ecran));
CN5:=TCheckBox(PGGetControl('CN5',Ecran));
CEtab:=TCheckBox(PGGetControl('CETAB',Ecran));
CL1:=TCheckBox(PGGetControl('CL1',Ecran));
CL2:=TCheckBox(PGGetControl('CL2',Ecran));
CL3:=TCheckBox(PGGetControl('CL3',Ecran));
CL4:=TCheckBox(PGGetControl('CL4',Ecran));
Rupture:=THEdit(PGGetControl('XX_RUPTURE1',Ecran));
Champ1:=THEdit(PGGetControl('XX_VARIABLE1',Ecran));
If (Champ1<>nil) and (Rupture<>nil) then
  Champ1.text:='';Rupture.Text:='';
if (CEtab<>nil) and (Champ1<>nil) and (Rupture<>nil) then
   if (CEtab.Checked=True) then
     begin Champ1.text:='Etablissement :'; Rupture.Text:=PrefTable+'_ETABLISSEMENT';  end;
if (CN1<>nil)and(CN2<>nil)and(CN3<>nil)and(CN4<>nil)and(CN5<>nil) then
 If (Champ1<>nil) and (Rupture<>nil) then
   begin
   if (CN1.Checked=True) then
       begin
       Champ1.text:=VH_Paie.PGLibelleOrgStat1;
       Rupture.Text:=PrefTable+'_TRAVAILN1'; //
       end;
   if (CN2.Checked=True) then
       begin
       Champ1.text:=VH_Paie.PGLibelleOrgStat2;
       Rupture.Text:=PrefTable+'_TRAVAILN2';
       end;
   if (CN3.Checked=True) then
       begin
       Champ1.text:=VH_Paie.PGLibelleOrgStat3;
       Rupture.Text:=PrefTable+'_TRAVAILN3';
       end;
   if (CN4.Checked=True) then
       begin
       Champ1.text:=VH_Paie.PGLibelleOrgStat4;
       Rupture.Text:=PrefTable+'_TRAVAILN4';
       end;
   if (CN5.Checked=True) then
       begin
       Champ1.text:=VH_Paie.PGLibCodeStat;
       Rupture.Text:=PrefTable+'_CODESTAT';
       end;
   end;
if (CL1<>nil)and(CL2<>nil)and(CL3<>nil)and(CL4<>nil) then
 If (Champ1<>nil) and (Rupture<>nil) then
   begin
   if (CL1.Checked=True) then
       begin
       Champ1.text:=VH_Paie.PgLibCombo1;
       Rupture.Text:=PrefTable+'_LIBREPCMB1'; //
       end;
   if (CL2.Checked=True) then
       begin
       Champ1.text:=VH_Paie.PgLibCombo2;
       Rupture.Text:=PrefTable+'_LIBREPCMB2';
       end;
   if (CL3.Checked=True) then
       begin
       Champ1.text:=VH_Paie.PgLibCombo3;
       Rupture.Text:=PrefTable+'_LIBREPCMB3';
       end;
   if (CL4.Checked=True) then
       begin
       Champ1.text:=VH_Paie.PgLibCombo4;
       Rupture.Text:=PrefTable+'_LIBREPCMB4';
       end;
   end;

{if (CN1.Checked=False)And(CN2.Checked=False)And(CN3.Checked=False)And(CN4.Checked=False)AND(CN5.Checked=False)AND(CETAB.Checked=False)  then
  begin
  Champ1.text:='';Rupture.Text:='';
  end;}
end;


function AffectSqlEtat(Ecran:TForm) : string ;
var
Alpha,CEtab,CTrav1,CTrav2,CTrav3,CTrav4,CTrav5:TCheckBox;
SQL,AddChampTri:String;
begin
Tip:=TFQRS1(Ecran).TypeEtat;
Nat:=TFQRS1(Ecran).NatureEtat;
Modele:=TFQRS1(Ecran).CodeEtat;
PrefTable := PGRendPrefixeFromQrs1 (Ecran,Tip,Nat,Modele);
  //Bulletin de Paie selon rupture associée
CTrav1:=TCheckBox(PGGetControl('CN1',Ecran));
CTrav2:=TCheckBox(PGGetControl('CN2',Ecran));
CTrav3:=TCheckBox(PGGetControl('CN3',Ecran));
CTrav4:=TCheckBox(PGGetControl('CN4',Ecran));
CTrav5:=TCheckBox(PGGetControl('CN5',Ecran));
CEtab:=TCheckBox(PGGetControl('CETAB',Ecran));
Alpha:=TCheckBox(PGGetControl('CALPHA',Ecran));

// TRI SUR MATRICULE SALARIE + CHAMPS RUPTURE
if CTrav1<>nil then
  if CTrav1.Checked=True then
  begin  AddChampTri:=PrefTable+'_TRAVAILN1,';  end;
if CTrav2<>nil then
  if CTrav2.Checked=True then
  begin AddChampTri:=PrefTable+'_TRAVAILN2,';   end;
if CTrav3<>nil then
   if CTrav3.Checked=True then
   begin AddChampTri:=PrefTable+'_TRAVAILN3,';  end;
if CTrav4<>nil then
   if CTrav4.Checked=True then
   begin  AddChampTri:=PrefTable+'_TRAVAILN4,';  end;
if CTrav5<>nil then
   if CTrav5.Checked=True then
   begin AddChampTri:=PrefTable+'_CODESTAT,';   end;
if Cetab<>nil then
   if Cetab.Checked=True then
   begin AddChampTri:=PrefTable+'_ETABLISSEMENT,';  End; //end;

 // TRI SUR LIBELLE SALARIE + CHAMPS RUPTURE    //Cas Combiné avec tri alphabétique

if Alpha<>nil then
   if Alpha.Checked=True then
   begin AddChampTri:='PPU_LIBELLE,PPU_PRENOM,';   end;
if  (Cetab.Checked=True) and (Alpha.Checked=True) then
   begin AddChampTri:=PrefTable+'_ETABLISSEMENT,PPU_LIBELLE,PPU_PRENOM,';  end;
if  (CTrav1.Checked=True) and (Alpha.Checked=True) then
   begin  AddChampTri:=PrefTable+'_TRAVAILN1,PPU_LIBELLE,PPU_PRENOM,';  end;
if  (CTrav2.Checked=True) and (Alpha.Checked=True) then
   begin AddChampTri:=PrefTable+'_TRAVAILN2,PPU_LIBELLE,PPU_PRENOM,';   end;
if  (CTrav3.Checked=True) and (Alpha.Checked=True) then
   begin AddChampTri:=PrefTable+'_TRAVAILN3,PPU_LIBELLE,PPU_PRENOM,';  end;
if  (CTrav4.Checked=True) and (Alpha.Checked=True) then
   begin AddChampTri:=PrefTable+'_TRAVAILN4,PPU_LIBELLE,PPU_PRENOM,';   end;
if  (CTrav5.Checked=True) and (Alpha.Checked=True) then
   begin AddChampTri:=PrefTable+'_CODESTAT,PPU_LIBELLE,PPU_PRENOM,';   end;


    //Bulletin de paie par défaut
if (CTrav1<>nil)and(CTrav2<>nil)and(CTrav3<>nil)and(CTrav4<>nil)and(CTrav5<>nil)and(Cetab<>nil)and(Alpha<>nil) then
if (Alpha.Checked=False)and(CTrav1.Checked=False)and(CTrav2.Checked=False)and(CTrav3.Checked=False)and(CTrav4.Checked=False)and(CTrav5.Checked=False)and(Cetab.Checked=False)then
  begin
  AddChampTri:='';
  end;
 SQL:='';
// SQL:=' '+PrefTable+'_DATEDEBUT,'+PrefTable+'_DATEFIN,'+AddChampTri+' '+PrefTable+'_SALARIE';//+PrefTable+'_ETABLISSEMENT';
 SQL:=' '+AddChampTri+' '+PrefTable+'_SALARIE';//+PrefTable+'_ETABLISSEMENT';
 RecupOrderBy:=SQL;
 result:=SQL;
end;

{function renvoyant le taux de conversion Devise/Euro à multiplier
2 cas possible :
 - Dossier tenue en euro : on multiplie le montant par le taux de parité ex 6.55957 Frs
 - Dossier non tenue en euro : on multiplie le montant par (1/parité) ex 1/6.55957=0.15..
-------------
Ce taux de conversion est utilisé pour la conversion des états (cf monnaie inversée).}
function  RendTauxConvertion : string;
var
QRech : TQuery;
begin
Result:='1';
if (VH_Paie.PGMonnaieTenue<>'') then
   Begin
   QRech:=OpenSql('SELECT D_PARITEEURO FROM DEVISE WHERE D_DEVISE="'+VH_Paie.PGMonnaieTenue+'"',True);
   if not QRech.eof then //PORTAGECWAS
     if QRech.Fields[0].AsFloat<>0 then
       Begin
       if (VH_Paie.PGTenueEuro=True) then Result:=FloatToStr(  QRech.Fields[0].AsFloat);
       if (VH_Paie.PGTenueEuro=False)then Result:=FloatToStr(1/QRech.Fields[0].AsFloat);
       End;
   Ferme(Qrech);
   End;
end;


function FormatCase(St, StType: string ; Esp : integer): string;
Const FormString = ['0'..'9','a'..'z','A'..'Z'] ;   //PT1 Ajout 'a'..'z','A'..'Z'
Var
StFormat: string;
i,j,Lg,DebHr : integer;
begin
if St='' then begin result:=''; exit; end;
StType:=AnsiUpperCase(StType);
Lg:=length(st);
j:=1;
StFormat:=Stringofchar (' ',Lg+5*(Lg-1));
If StType='STR' then
  For i:=1 to Lg do
  if St[i] in FormString then
    begin
    StFormat[j]:=St[i];
    j:=j+Esp;
    end;

If StType='NBR' then
  For i:=1 to Lg do
  begin
  StFormat[j]:=St[i];
  j:=j+Esp;
  end;

If StType='DATE' then
  For i:=1 to Lg do
    if (St[i]<>'/') then// and (i<>7) and (i<>8) then
      begin
      StFormat[j]:=St[i];
      j:=j+Esp;
      if (i=2) or (i=5) then j:=j+2;   //PT4 ajout espacement
      end;

If StType='HR' then
  Begin
  DebHr:=Pos(':',st)-2;
  if DebHr>0 then
   Begin
    St:=Copy(st,DebHr,lg);
    StFormat:=Stringofchar (' ',Lg+5*(Lg));
    For i:=1 to Lg do
      if (St[i]<>':') and (i<6) then
        begin
        StFormat[j]:=St[i];
        if i=2 then Esp:=ESp+1;
        if i=4 then Esp:=ESp-1;
        j:=j+Esp;
        end;
   End
   else StFormat:='';
  End;

If StType='SS' then
  For i:=1 to Lg do
    if St[i] in FormString then
      begin
      StFormat[j]:=St[i];
      if (i=1) or (i=5)then j:=j+1;
      if (i=3)or (i=7) or (i=10) or (i=13) then j:=j+2; //PT4 ajout espacement
      j:=j+Esp;
      end;

result:=Trim(StFormat);
end;

function RemplaceVirguleParPoint(St: String): String;
var
i : Integer;
begin
I:=Pos(',',St);
If St[i]=',' then St[i]:='.';
//For i:=1 to Length(St) do
  // If St[i]=',' then St[i]:='.';
result := st;
end;

procedure VisibiliteChampLibre(Ecran : TForm);
var
num : integer;
numero,libelle,Min,Max : string;
Edit : THEdit;
TLieu,Lbl : THLabel;
Check : TCheckBox;
Begin
if Ecran is TFQRS1 then //PT3 PrefTable forcé à PPU si fiche <> QRS1
  Begin
  Tip:=TFQRS1(Ecran).TypeEtat;
  Nat:=TFQRS1(Ecran).NatureEtat;
  Modele:=TFQRS1(Ecran).CodeEtat;
  PrefTable := PGRendPrefixeFromQrs1 (Ecran,Tip,Nat,Modele);
// PT6 05/04/2004 PH V_50 Prise en compte des états des provisions cp et Rtt
  if (nat='PCG') and ((Modele='PRT') OR (Modele='PDP') OR (Modele='PRP')) then PrefTable:='PSA';
  end
else PrefTable:='PPU';
For num := 1 to VH_Paie.PgNbCombo do
  begin
  Numero:=InttoStr(num);
  if Num > VH_Paie.PgNbCombo then break;
  Libelle:='';
  if Num = 1 then libelle := VH_Paie.PgLibCombo1;
  if Num = 2 then libelle := VH_Paie.PgLibCombo2;
  if Num = 3 then libelle := VH_Paie.PgLibCombo3;
  if Num = 4 then libelle := VH_Paie.PgLibCombo4;
  if (Libelle<>'') then
    Begin
   // RecupMinMaxTablette('CC','','PL'+Numero,Min,Max);
    Edit:=THEdit(PGGetControl(PrefTable+'_LIBREPCMB'+Numero,Ecran));
    if Edit <> NIL then Begin Edit.Text:=Min; Edit.Visible :=TRUE; End;
    Edit:=THEdit(PGGetControl(PrefTable+'_LIBREPCMB'+Numero+'_',Ecran));
    if Edit <> NIL then Begin Edit.Text:=Max; Edit.Visible :=TRUE; End;
    Check:=TCheckBox(PGGetControl('CL'+Numero,Ecran));
    if Check <> NIL then Check.Visible :=TRUE;
    TLieu:=THLabel(PGGetControl('T'+PrefTable+'_LIBREPCMB'+Numero,Ecran));
    if (TLieu <> NIL)   then
      begin
      TLieu.Visible :=TRUE;
      if Num = 1 then Begin TLieu.Caption :=VH_Paie.PgLibCombo1+' de'; end;
      if Num = 2 then Begin TLieu.Caption :=VH_Paie.PgLibCombo2+' de'; end;
      if Num = 3 then Begin TLieu.Caption :=VH_Paie.PgLibCombo3+' de'; end;
      if Num = 4 then Begin TLieu.Caption :=VH_Paie.PgLibCombo4+' de'; end;
      end;
      Lbl:=THLabel(PGGetControl (('TLA'+Numero),Ecran));
      if Lbl <> NIL then Lbl.visible:=True;
    End;
  end;
End;

procedure BloqueChampLibre(Ecran:TForm);
var
TabLieuTravail : array[1..10] of TCheckBox;
PosCheck,PosUnCheck,i,l : integer;
Ok : boolean ;
BEGIN
if Ecran is TFQRS1 then //PT3 PrefTable forcé à PPU si fiche <> QRS1
  Begin
  Tip:=TFQRS1(Ecran).TypeEtat;
  Nat:=TFQRS1(Ecran).NatureEtat;
  Modele:=TFQRS1(Ecran).CodeEtat;
  PrefTable := PGRendPrefixeFromQrs1 (Ecran,Tip,Nat,Modele);
  End
Else  PrefTable:='PPU';
TabLieuTravail[1]:=TCheckBox(PGGetControl('CN1',Ecran));
TabLieuTravail[2]:=TCheckBox(PGGetControl('CN2',Ecran));
TabLieuTravail[3]:=TCheckBox(PGGetControl('CN3',Ecran));
TabLieuTravail[4]:=TCheckBox(PGGetControl('CN4',Ecran));
TabLieuTravail[5]:=TCheckBox(PGGetControl('CN5',Ecran));
TabLieuTravail[6]:=TCheckBox(PGGetControl('CETAB',Ecran));
TabLieuTravail[7]:=TCheckBox(PGGetControl('CL1',Ecran));
TabLieuTravail[8]:=TCheckBox(PGGetControl('CL2',Ecran));
TabLieuTravail[9]:=TCheckBox(PGGetControl('CL3',Ecran));
TabLieuTravail[10]:=TCheckBox(PGGetControl('CL4',Ecran));
PosUnCheck:=0;
PosCheck:=0;
//if (TabLieuTravail[1]<>nil) and (TabLieuTravail[2]<>nil) then
//if (TabLieuTravail[3]<>nil) and (TabLieuTravail[4]<>nil) and (TabLieuTravail[5]<>nil) then
For l:=1 to 10 do
  if (TabLieuTravail[l]<>nil) then Ok:=False else Begin Ok:=True; break; End;

If Ok=False then
  begin
  //Coche une rupture
  For i:=1 to 10 do
    if (TabLieuTravail[i].checked=True) then PosCheck:=i;
  If PosCheck>0 then
    For i:=1 to 10 do
      Begin
      if i<>PosCheck then
        Case i of        //rend unenable(false) les autres champs de rupture
        1..4 : Begin
               TabLieuTravail[i].enabled:=False;
               {ChampLieu:=THEdit(PGGetControl((PrefTable+'_TRAVAILN'+IntToStr(i)),Ecran));
               if ChampLieu <> NIL then begin ChampLieu.Enabled :=False; ChampLieu.Text:=''; end;
               ChampLieu_:=THEdit(PGGetControl((PrefTable+'_TRAVAILN'+IntToStr(i)+'_'),Ecran));
               if ChampLieu_ <> NIL then begin ChampLieu_.enabled :=False;ChampLieu_.Text:=''; end;
               }End;
           5 : begin
               TabLieuTravail[i].enabled:=False;
               {ChampLieu:=THEdit(PGGetControl(PrefTable+'_CODESTAT',Ecran));
               if ChampLieu <> NIL then begin ChampLieu.Enabled :=False; ChampLieu.Text:=''; end;
               ChampLieu_:=THEdit(PGGetControl(PrefTable+'_CODESTAT_',Ecran));
               if ChampLieu_ <> NIL then begin ChampLieu_.enabled :=False; ChampLieu_.Text:=''; end;
               }end;
           6 : begin
               TabLieuTravail[i].enabled:=False;
               {ChampLieu:=THEdit(PGGetControl(PrefTable+'_ETABLISSEMENT',Ecran));
               if ChampLieu <> NIL then begin ChampLieu.Enabled :=False; ChampLieu.Text:=''; end;
               ChampLieu_:=THEdit(PGGetControl(PrefTable+'_ETABLISSEMENT_',Ecran));
               if ChampLieu_ <> NIL then begin ChampLieu_.enabled :=False;  ChampLieu_.Text:=''; end;
               }end;
       7..10 : Begin
               TabLieuTravail[i].enabled:=False;
               {l:=i-6;
               ChampLieu:=THEdit(PGGetControl(PrefTable+'_LIBREPCMB'+IntToStr(l),Ecran));
               if ChampLieu <> NIL then begin ChampLieu.Enabled :=False; ChampLieu.Text:=''; end;
               ChampLieu_:=THEdit(PGGetControl(PrefTable+'_LIBREPCMB'+IntToStr(l)+'_',Ecran));
               if ChampLieu_ <> NIL then begin ChampLieu_.enabled :=False;  ChampLieu_.Text:=''; end;
               }End;
        End; //Fin Case
      end; //Fin For

     //Décoche une rupture ,  rend enable(True) les autres champs de rupture
  For i:=1 to 10 do
    if (TabLieuTravail[i].checked=False) and (TabLieuTravail[i].enabled=True) then PosUnCheck:=i;
  If (PosCheck=0) and (PosUnCheck>0) then
  For i:=1 to 10 do
    begin
      TabLieuTravail[i].enabled:=True;
      {Case i of
       1..4 : Begin
              If i<=VH_Paie.PGNbreStatOrg then
                begin
                Min:=''; Max:='';
                if (i = 1) and (VH_Paie.PGLibelleOrgStat1<>'') then RecupMinMaxTablette('CC','','PAG',Min,Max);
                if (i = 2) and (VH_Paie.PGLibelleOrgStat2<>'') then RecupMinMaxTablette('CC','','PST',Min,Max);
                if (i = 3) and (VH_Paie.PGLibelleOrgStat3<>'') then RecupMinMaxTablette('CC','','PUN',Min,Max);
                if (i = 4) and (VH_Paie.PGLibelleOrgStat4<>'') then RecupMinMaxTablette('CC','','PSV',Min,Max);
                ChampLieu:=THEdit(PGGetControl((PrefTable+'_TRAVAILN'+IntToStr(i)),Ecran));
                if ChampLieu <> NIL then Begin ChampLieu.Enabled :=True; ChampLieu.text:=Min; end;
                ChampLieu_:=THEdit(PGGetControl((PrefTable+'_TRAVAILN'+IntToStr(i)+'_'),Ecran));
                if ChampLieu_ <> NIL then Begin ChampLieu_.enabled :=True;  ChampLieu_.text:=Max; end;
                end;
              End;
          5 : begin
              If (VH_Paie.PGLibCodeStat<>'') then
                Begin
                Min:=''; Max:='';
                RecupMinMaxTablette('CC','','PSQ',Min,Max);
                ChampLieu:=THEdit(PGGetControl(PrefTable+'_CODESTAT',Ecran));
                if ChampLieu <> NIL then Begin ChampLieu.Enabled :=True;ChampLieu.text:=Min; end;
                ChampLieu_:=THEdit(PGGetControl(PrefTable+'_CODESTAT_',Ecran));
                if ChampLieu_ <> NIL then Begin ChampLieu_.enabled :=True; ChampLieu_.text:=Max; End;
                end;
              end;
          6 : begin
              ChampLieu:=THEdit(PGGetControl(PrefTable+'_ETABLISSEMENT',Ecran));
              RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
              if ChampLieu <> NIL then begin ChampLieu.Enabled :=True; ChampLieu.text:=Min; End;
              ChampLieu_:=THEdit(PGGetControl(PrefTable+'_ETABLISSEMENT_',Ecran));
              if ChampLieu_ <> NIL then begin ChampLieu_.enabled :=True; ChampLieu_.text:=Max; end;
              end;
      7..10 : Begin
              l:=i-6;  Min:=''; Max:='';  Libelle:='';
              if l= 1 then libelle := VH_Paie.PgLibCombo1;  if l= 2 then libelle := VH_Paie.PgLibCombo2;
              if l= 3 then libelle := VH_Paie.PgLibCombo3;  if l= 4 then libelle := VH_Paie.PgLibCombo4;
             // if libelle<>'' then RecupMinMaxTablette('CC','','PL'+IntToStr(l),Min,Max);
              ChampLieu:=THEdit(PGGetControl((PrefTable+'_LIBREPCMB'+IntToStr(l)),Ecran));
              if ChampLieu <> NIL then Begin ChampLieu.Enabled :=True; ChampLieu.text:=Min; end;
              ChampLieu_:=THEdit(PGGetControl((PrefTable+'_LIBREPCMB'+IntToStr(l)+'_'),Ecran));
              if ChampLieu_ <> NIL then Begin ChampLieu_.enabled :=True;  ChampLieu_.text:=Max; end;
              End;
       end; } //Fin Case
    end;
  end; // End Ok=True
End;

// Fonction qui rend le prefixe de la table principale dans la cas de fiche QRS1
function  PGRendPrefixeFromQrs1 (Ecran:TForm; Tip,Nat,Modele : String): string;
var TableEtat : String;
begin
result := '';
if (Ecran.Name = 'EDITBUL_ETAT') OR (Ecran.Name = 'EDITBULSAL_ETAT') then result:='PPU'
else
  begin
  TableEtat:=GetTableFromEtat(Tip,Nat,Modele);
  result:=TableToPrefixe(TableEtat) ;
  end;
end;

//DEB PT2
Procedure AffectCritereAlpha(Ecran : TForm ;TriAlpha : Boolean; Champ1,Champ2 : String);
var min,max : string;
Begin
exit; { PT5 pour désactiver l'affectation alphbétique }
if TriAlpha then
   Begin
   THEdit(PGGetControl(Champ1,Ecran)).DataType:='PGSALARIE1';
   THEdit(PGGetControl(Champ1+'_',Ecran)).DataType:='PGSALARIE1';
   THEdit(PGGetControl(Champ1,Ecran)).Name:=Champ2;
   THEdit(PGGetControl(Champ1+'_',Ecran)).Name:=Champ2+'_';
   RecupMinMaxTablette('PG','SALARIES','PSA_LIBELLE',Min,Max);
   THEdit(PGGetControl(Champ2,Ecran)).Text:=Min;
   THEdit(PGGetControl(Champ2+'_',Ecran)).Text:=Max;
   End
  else
   Begin
   THEdit(PGGetControl(Champ2,Ecran)).Name:=Champ1;
   THEdit(PGGetControl(Champ2+'_',Ecran)).Name:=Champ1+'_';
   THEdit(PGGetControl(Champ1,Ecran)).DataType:='PGSALARIE';
   THEdit(PGGetControl(Champ1+'_',Ecran)).DataType:='PGSALARIE';
   RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
   THEdit(PGGetControl(Champ1,Ecran)).Text:=Min;
   THEdit(PGGetControl(Champ1+'_',Ecran)).Text:=Max;
   End;
End;
//FIN PT2

{!!!!Utilisé pour test développement états chaînés
Function de retour des requêtes de lancement des états
L'edition en masse est basée sur des appels de LanceEtat passant en paramètre le SQL de l'état..
Afin de toujours utiliser la même requête en edition QRS1 ou en masse..
Ces dernières sont regoupés dans cette fonction commune d'appel}
function  RendRequeteSQLEtat(NomEtat,Champ,Critere,OrderBy : string) : string;
Begin
Result:='';
if NomEtat='BULLETIN' then
  Begin
  Result:='SELECT PPU_ETABLISSEMENT,PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN, '+
       'PPU_EDTDEBUT,PPU_EDTFIN,PPU_DATEENTREE,PPU_DATESORTIE,PPU_CIVILITE,'+
       'PPU_LIBELLE,PPU_PRENOM,PPU_AUXILIAIRE,'+
       'PPU_ADRESSE1,PPU_ADRESSE2,PPU_ADRESSE3,PPU_CODEPOSTAL,PPU_VILLE,'+
       'PPU_NUMEROSS,PPU_CONVENTION,PPU_CODEEMPLOI,PPU_INDICE,PPU_NIVEAU,'+
       'PPU_QUALIFICATION,PPU_COEFFICIENT,PPU_PGMODEREGLE,'+
       'PPU_TRAVAILN1,PPU_TRAVAILN2,PPU_TRAVAILN3,PPU_TRAVAILN4,PPU_CODESTAT,'+
       'PPU_LIBELLEEMPLOI,PPU_PAYELE,PPU_CHEURESTRAV,PPU_CBRUT,PPU_CNETAPAYER,'+
       'PPU_CBRUTFISCAL,PPU_CNETIMPOSAB,PPU_CCOUTPATRON,PPU_CCOUTSALARIE,PPU_DATEANCIENNETE,'+
       'PHB_NATURERUB,PHB_RUBRIQUE,PHB_LIBELLE,PHB_BASEREM,PHB_TAUXREM,PHB_COEFFREM,'+
       'PHB_MTREM,PHB_IMPRIMABLE,PHB_BASECOT,PHB_ORDREETAT,PHB_TAUXSALARIAL,'+
       'PHB_MTSALARIAL,PHB_TAUXPATRONAL,PHB_MTPATRONAL,PHB_BASEREMIMPRIM,'+
       'PHB_TAUXREMIMPRIM,PHB_COEFFREMIMPRIM,PHB_BASECOTIMPRIM,PHB_TAUXSALIMPRIM,'+
       'PHB_TAUXPATIMPRIM,PHB_ORGANISME,PHB_SENSBUL, '+
       'ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_CODEPOSTAL,ET_VILLE,ET_SIRET,ET_APE '+
       'FROM PAIEENCOURS '+
       'LEFT JOIN HISTOBULLETIN ON PPU_ETABLISSEMENT=PHB_ETABLISSEMENT AND '+
       'PPU_SALARIE=PHB_SALARIE AND PPU_DATEDEBUT=PHB_DATEDEBUT AND PPU_DATEFIN=PHB_DATEFIN '+
       'LEFT JOIN ETABLISS ON ET_ETABLISSEMENT=PPU_ETABLISSEMENT '+
       'WHERE PHB_NATURERUB<>"BAS" '+
       'AND (PHB_MTSALARIAL<>0 OR PHB_MTPATRONAL<>0 OR PHB_MTREM<>0 OR PHB_RUBRIQUE like "%.%" OR PHB_BASEREM<>0) '+
       'AND PHB_IMPRIMABLE="X" AND PHB_ORDREETAT>0 '+Critere+OrderBy;

  End
else
  if  NomEtat='JOURNALPAIE' then
    begin
    Result:= 'SELECT DISTINCT PPU_SALARIE,PSA_LIBELLE,PSA_PRENOM,'+Champ+
             'PSA_DATEENTREE,PSA_DATESORTIE FROM PAIEENCOURS '+
             'LEFT JOIN SALARIES ON PPU_SALARIE=PSA_SALARIE '+
             critere+OrderBy;
    End;

End;

end.
