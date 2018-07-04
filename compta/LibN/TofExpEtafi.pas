{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 10/02/2003
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit TofExpEtafi;

interface
uses  Windows,Controls,StdCtrls,ExtCtrls,Graphics,Grids,
      Classes,ComCtrls,sysutils,
      {$IFNDEF EAGLCLIENT}
      {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
      {$ENDIF}
      {$IFDEF VER150}
      Variants,
      {$ENDIF}
      Ent1,Vierge,
      HCtrls,HEnt1,LettUtil,HMsgBox,LettAuto,UTOF,UTOB,
      paramsoc,forms,menus,{filtre, SG6 Gestion des filtres FQ 14826 15/11/04}
      CPLISGEN_TOF, // CP_LanceFicheLisGen
      HTB97,UObjFiltres {SG6 15/11/04 FQ 14826} ;

procedure CP_LanceFicheExpEtafi;

Type
     TOF_ExpEtafi = class (TOF)
     private
       ObjFiltre : TObjFiltre; //Gestion des Filtres SG6 1511/04  FQ 14826
       F6,UnEtab : boolean ;
       DecLocal : integer ;
       CptAGen : TToolBarButton97 ;
       DateDebut,DateFin : TDateTime ;
       CFormat,Cexo,CEtab : THValComboBox ;
{$IFNDEF SPEC350}
       CQualif: THValComboBox ;
{$ELSE}
       Simu : boolean ;
{$ENDIF}
       DatDeb,DatFin,FileName,TBenef,TPerte,Tresult : THedit ;
       RMonnaie,RSeparateur : TRadioGroup ;
       PControl : TPageControl ;
       CHuit,CCompensation : TCheckBox ;
       GenEtab : TCheckBox ;
       TheTob : TOB ;
       TobA,TobB,TobC,TobC2,TobG :Tob ;
       // TOBA : Liste des comptes auxiliaires
       // TOBB : Liste des écritures dont les comptes généraux ne sont pas collectifs
       // TOBC : Liste des écritures dont les comptes généraux sont collectifs
       // TOBC2 : Liste des écritures dont les comptes généraux sont collectifs
       // TOBG : Liste des comptes généraux (Compte + Nature + Libellés)
       // TheTob : Liste des écritures qui seront exportés au final dans le fichier
       TobGen : Tob ;
       function  CreeLaFille : TOB ;
       function  ParamOK : boolean ;
       procedure ExportExo(Exo : string) ;
       function  RecupTypeEcriture: string ;
       procedure RecupSolde(Exo : string;N : integer) ;
       procedure ExportLaTob(Fname : string ) ;
       procedure OnChangeCExo(Sender: TObject) ;
       function  CalculExo(Exo : string; N : integer) : string;
       function  FormatSpec(Dbl : double;Typ,Deci : Integer) : string ;
       function  RecupLib(Compte : string ) : string ;
       function  CreeLeSelect(Exo : string;N : integer): string ;
       function RemplitTob(General,Aux,Libelle,EcrAn: string;N: integer; Debit,Credit: double) : Tob;
       procedure EnableButtons ;
       procedure OnChangeCFormat(Sender: TObject) ;
       procedure ChargeTobG;
       procedure ChargeTobA;
       procedure ChargeTobC(Exo: String ; N : Integer);
       function  RecupLibAux(Auxi: string): string;
       function  CreeLigne: TOB;
       procedure CreeLigneTobCompensee(Gene, Etab, EAN, prefix: String; Deb,Cre: Double);
       function  FormateLeGeneral(Gene, Etab, Sens: String): String;
       procedure CptAgenOnClick(Sender: TObject);
       procedure ChargeLaTobGen;
       procedure GenEtabOnClick (Sender : TObject) ;
       {procedure RetraiteSolde(TobOK : Tob; const General, Exo : String); JP 18/07/05 : FQ 13318 : devenu inutile}
       procedure SetCredDeb(aTob : TOB); {JP 08/07/05 : FQ 13318}
     public
       FNomFiltre : String ;
       procedure OnLoad ; override ;
       procedure OnUpdate ; override ;
       procedure OnClose ; override ;
       procedure OnArgument(S: string); override; {SG6 Gestion des filtres FQ 14826 15/11/04}
       {JP 28/06/06 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc}
       procedure GereEtablissement;
       {JP 28/06/06 : FQ 16149 : on s'assure que le filtre coincide avec les restrictions utilisateurs sur l'établissement}
       procedure ControlEtab;
     end;

implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  {$IFDEF EAGLCLIENT}
  MaineAGL;
  {$ELSE}
  Fe_Main;
  {$ENDIF}

procedure CP_LanceFicheExpEtafi;
begin
  AGLLanceFiche('CP', 'CPEXPETAFI', '', '', '') ; // ETAFI Servant
end;

Function TOF_ExpEtafi.FormatSpec(Dbl : double;Typ,Deci : Integer) : string ;
var  OldDecimalSeparator : char ;  i,Expo,Taille : integer ;
begin
Expo:=1 ;for i:=1 to Deci do Expo:=Expo*10;
if F6 then Taille:=17 else Taille:=16 ;
OldDecimalSeparator:=DecimalSeparator ;
case Typ of
  0 : begin DecimalSeparator:=',' ; Result:=format('%*.2f',[Taille,Dbl]) ; end ;
  1 : begin DecimalSeparator:='.' ; Result:=format('%*.2f',[Taille,Dbl]) ; end ;
  2 : Result:=format('%16.0f',[Dbl*Expo]) ;
  else
  Result:='' ;
  end ;
DecimalSeparator:=OldDecimalSeparator ;
End ;


{***********A.G.L.***********************************************
Auteur  ...... : Stephane Guillon
Créé le ...... : 15/11/2004
Modifié le ... :   /  /    
Description .. : Evt On argument 
Suite ........ : - Gestion des filtres SG6 15/11/04 FQ 14826
Mots clefs ... : ARGUMENT,FILTRE
*****************************************************************}
procedure TOF_ExpEtafi.OnArgument(S: string);
var
  Composants : TControlFiltre; //SG6   Gestion des Filtes 15/11/04   FQ 14826
  FFiltres:THValComboBox;
  POPF:TPopupMenu;
  BFILTRE:TToolBarButton97;
  Pages:TPageControl;
begin
  //SG6 10/11/04 Gestion des Filtres FQ 14826
  FFILTRES:=THValComboBox(GetControl('FFILTRES')); if FFILTRES=nil then Exit ;
  POPF:=TPopupMenu(GetControl('POPF')); if POPF=nil then Exit;
  BFILTRE:=TToolBarButton97(GetControl('BFILTRES')); if BFILTRE=nil then Exit;
  Pages:=TPageControl(GetControl('PCONTROL')); if Pages=nil then Exit;
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, 'TOFEXPETAFI');
end;


function TOF_ExpEtafi.RecupLib(Compte : string ) : string ;
var TobL: Tob ;
begin
Result:='' ;
TobL:=TobG.Findfirst(['G_GENERAL'],[Compte],False);
if TobL <> Nil then Result:=TobL.GetValue('G_LIBELLE') ;
end ;

function TOF_ExpEtafi.RecupLibAux(Auxi : string ) : string ;
var TobL: Tob ;
begin
Result:='' ;
TobL:=TobA.Findfirst(['T_AUXILIAIRE'],[Auxi],False);
if TobL <> Nil then Result:=TobL.GetValue('T_LIBELLE') ;
end ;

{function TOF_ExpEtafi.RecupLib(Compte : string ) : string ;
var Q : TQuery ;
begin
Q:=OpenSql('SELECT G_LIBELLE FROM GENERAUX WHERE G_GENERAL="'+Compte+'"' ,True) ;
if not Q.Eof then Result:=Q.Fields[0].AsString else Result:='' ;
ferme(Q) ;
end ;}

procedure TOF_ExpEtafi.EnableButtons ;
begin

    Ecran.HelpContext:=999999429;

if (CFormat.ItemIndex=0) then CCompensation.Enabled:=TRUE else CCompensation.Enabled:=FALSE ;
end ;

procedure TOF_ExpEtafi.OnLoad ;
var
  zz: Integer ;
begin
  {JP 18/07/05 : à faire dans le OnLoad et non pas dans le OnArgument comme c'était le cas jusqu'à présent}
  ObjFiltre.Charger;

CFormat:=THValComboBox(GetControl('CFORMAT')) ;
CExo   :=THValComboBox(GetControl('CEXO')) ;
CEtab  :=THValComboBox(GetControl('CETAB')) ;
DatFin  :=THEdit(GetControl('DATEFIN')) ;
DatDeb  :=THEdit(GetControl('DATEDEB')) ;
FileName:=THEdit(GetControl('FILENAME')) ;
TBenef  :=THEdit(GetControl('TBENEF')) ;
TPerte  :=THEdit(GetControl('TPERTE')) ;
CptAGen:=TToolBarButton97(GetControl('CPTAGEN')) ;
TResult :=THEdit(GetControl('TRESULT')) ;
RMonnaie   :=TRadioGroup(GetControl('RMONNAIE')) ;
RMonnaie.Visible := False;
RSeparateur:=TRadioGroup(GetControl('RSEPARATEUR')) ;
CHuit:=TCheckBox(GetControl('CHUIT')) ;
CCompensation:=TCheckBox(GetControl('CCOMPENSATION')) ;
GenEtab:=TCheckBox(GetControl('GENETAB')) ;
GenEtab.OnClick :=GenEtabOnClick ;
GenEtab.Enabled := GetParamSocSecur('SO_ETABLISCPTA',False);
PControl:=TPageControl(GetControl('PCONTROL')) ;
if (CFormat<>nil) and (not Assigned(CFormat.OnChange)) then CFormat.OnChange:=OnChangeCFormat ;
if (CExo<>nil) and (not Assigned(CExo.OnChange)) then CExo.OnChange:=OnChangeCexo ;
if (CptAGen<>nil) and (not Assigned(CptAGEn.OnClick)) then CptAGen.OnClick:=CptAgenOnClick ;
CFormat.ItemIndex:=0 ;

{JP 01/07/06 : FQ 16149 : nouvelle gestion des établissements}
GereEtablissement;
ObjFiltre.ApresChangementFiltre := ControlEtab;
{JP 01/07/06 : Si on a mis un filtre on conserve sa valeur, non ?}
if THValComboBox(GetControl('FFILTRES')).Value = '' then
  CExo.ItemIndex:=CExo.Items.Count-1 ;
OnChangeCExo(Cexo) ;

   TPerte.Text:=GetParamSocSecur('SO_OUVREPERTE','') ;
   TBenef.Text:=GetParamSocSecur('SO_OUVREBEN','') ;
   TResult.Text:=GetParamSocSecur('SO_OUVREBEN','') ;

DateDebut:=StrToDate(StDate1900) ;DateFin:=StrToDate(StDate2099) ;
FileName.Text:=ExtractFilePath(Application.ExeName)+'ExpEtafi.txt' ;
if (PControl<>nil) then PControl.ActivePage:=Pcontrol.Pages[0];
{$IFNDEF SPEC350}
CQualif:=THValComboBox(GetControl('CQUALIF')) ;
CQualif.ItemIndex:=0 ;
if V_PGI.LaSerie=S5 then
  begin
  zz:=CQualif.Values.IndexOf('PRE') ;
  if (zz<>-1) then begin CQualif.Values.Delete(zz) ; CQualif.Items.Delete(zz) ; end ;
  zz:=CQualif.Values.IndexOf('TOU') ;
  if (zz<>-1) then begin CQualif.Values.Delete(zz) ; CQualif.Items.Delete(zz) ; end ;
  end;
{$ENDIF}
EnableButtons ;
// InitFiltre SG6
end ;

procedure TOF_ExpEtafi.CptAgenOnClick(Sender: TObject) ;
var i : Integer ;
    St : String ;
begin
if TobGen=nil then
  begin
  TobGen:=Tob.Create('Gen',nil,-1) ;
  end ;
// recup du pointer de tobgen pour le passer a CPLISGEN_TOF
i:=Integer(TobGen) ; St:=IntToStr(i) ;
CP_LanceFicheLisGen(St);
end ;

procedure TOF_ExpEtafi.OnUpdate ;
begin
if not ParamOK then Exit ;
TheTob:=TOB.Create('_ETAFI',nil,-1) ;
ExportExo(Cexo.Value) ;
if TheTob<>nil then TheTob.Free ;
HShowMessage(TraduireMemoire('1;Export Etafi;L''export de la balance a été effectué;A;O;O;O'),'','')
end ;

procedure TOF_ExpEtafi.ChargeLaTobGen ;
var Q:TQuery ;
begin
if TobGen=nil then tobGen:=Tob.create('Gen',Nil,-1);
Q:=OpenSql('SELECT G_GENERAL, G_LIBELLE FROM GENERAUX', True) ;
TobGen.LoadDetailDb('Gen','','',Q,True) ;
Ferme(Q) ;
end ;

procedure TOF_ExpEtafi.ExportExo(Exo : string) ;
var i : integer ;
begin
{$IFDEF SPEC350}
Simu:=(HShowMessage(TraduireMemoire('1;Export Etafi;Voulez-vous inclure les écritures de simulations ?;Q;YN;N;N'),'','')=mrYes) ;
{$ENDIF}
if {(genEtab.Checked) and} (TobGen=Nil) then ChargeLaTobGen ;
{On traite les exercices de n (i = 0) à n - 3 (i = 3), s'ils existent}
for i:=0 to 3 do RecupSolde(Exo,i) ;
ExportLaTob(FileName.Text) ;
end ;

procedure TOF_ExpEtafi.OnClose ;
begin
FreeAndNil(ObjFiltre); // SG6 15/11/04 Gestion des Filtres FQ 14826
if TobGen<>nil then TobGen.Free ; TobGen:=Nil ;
end ;

function TOF_ExpEtafi.CreeLaFille : TOB ;
var T1 : TOB ;
begin
T1:=Tob.Create('_ETAFI2',TheTob,-1) ;
T1.AddChampSup('SGENERAL', False) ; T1.PutValue('SGENERAL', '') ;
T1.AddChampSup('GENERAL', False) ; T1.PutValue('GENERAL', '') ;
if F6 then T1.AddChampSup('AUXILIAIRE', False) ;
T1.AddChampSup('LIBELLE', False) ;
T1.AddChampSup('SOLDEAN', False) ; T1.PutValue('SOLDEAN', arrondi(0,DecLocal)) ;
T1.AddChampSup('MVDEBIT', False) ; T1.PutValue('MVDEBIT', arrondi(0,DecLocal));
T1.AddChampSup('MVCREDIT',False) ; T1.PutValue('MVCREDIT',arrondi(0,DecLocal));
T1.AddChampSup('SOLDE',   False) ; T1.PutValue('SOLDE',   arrondi(0,DecLocal));
T1.AddChampSup('SOLDEN1', False) ; T1.PutValue('SOLDEN1', arrondi(0,DecLocal));
T1.AddChampSup('SOLDEN2', False) ; T1.PutValue('SOLDEN2', arrondi(0,DecLocal));
T1.AddChampSup('SOLDEN3', False) ; T1.PutValue('SOLDEN3', arrondi(0,DecLocal));
Result:=T1 ;
end;

function TOF_ExpEtafi.ParamOK : boolean ;
var TMsg : array[1..10] of string ;
CarBour : string ;
Verif : boolean ;
begin
Verif := False;
TMsg[1]:='1;Export Etafi;La date est en dehors de l''exercice.;W;O;O;O' ;
TMsg[2]:='1;Export Etafi;La date est supérieure à la date inférieure.;W;O;O;O' ;
TMsg[3]:='1;Export Etafi;Le code "Etablissement principal" n''est pas renseigné.;W;O;O;O' ; // GV 21/05/02
ParamOK:=True ;
if StrToDate(DatDeb.Text)<DateDebut then begin HShowMessage(TraduireMemoire(TMsg[1]),'','') ; ParamOK:=False; DatDeb.SetFocus ; Exit ; end ;
if StrToDate(DatFin.Text)>DateFin   then begin HShowMessage(TraduireMemoire(TMsg[1]),'','') ; ParamOK:=False; DatFin.SetFocus ; Exit ; end ;
if StrToDate(DatDeb.Text)>StrToDate(DatFin.Text) then begin HShowMessage(TraduireMemoire(TMsg[2]),'','') ; ParamOK:=False; DatFin.SetFocus ; Exit ; end ;
//if ((GenEtab.Checked) and (SocMere.Text='')) then  begin HShowMessage(TraduireMemoire(TMsg[3]),'','') ; ParamOK:=False; SocMere.SetFocus ; Exit ; end ;
DecLocal:=V_PGI.OkDecV;
if CFormat.ItemIndex=1 then F6:=True else F6:=False ;
//if (CFormat.ItemIndex=0) and (not CCompensation.Checked) then F6:=TRUE ; // gv
if CEtab.ItemIndex<>0 then UnEtab:=True else UnEtab:=False ;

//RR 02/06/2005
//Rajout du test du caractère de bourrage
//Si<>0 on affiche des messages d'avertissement sur les généraux et/ou auxiliaires
//FQ 11774
CarBour := GetParamSocSecur('SO_BOURREGEN','') ;
if CarBour<>'0' then Verif := True ;

if Verif then
Begin
  if ExisteSQL('SELECT G_GENERAL FROM GENERAUX where SUBSTRING(G_GENERAL, LEN(G_GENERAL), 1)="' + CarBour +'"') then
  begin
   if HshowMessage(TraduireMemoire('1;Export ETAFI;Votre base contient des généraux avec des caractères de bourrage différent de "0", voulez-vous continuez ?;W;YN;Y;Y'),'','')<>mrYes then
      begin
      ParamOK:=False;
      if (PControl<>nil) then PControl.ActivePage:=Pcontrol.Pages[0];
      Exit ;
      end ;
  end;
end ;

// mode auxiliaire
if F6 then
BEGIN
  CarBour := GetParamSocSecur('SO_BOURREAUX','') ;
  if CarBour<>'0' then Verif := True ;

  if Verif then
  Begin
    if ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS where SUBSTRING(T_AUXILIAIRE, LEN(T_AUXILIAIRE), 1)="' + CarBour +'"') then
    begin
     if HshowMessage(TraduireMemoire('1;Export ETAFI;Votre base contient des auxiliaires avec des caractères de bourrage différent de "0", voulez-vous continuez ?;W;YN;Y;Y'),'','')<>mrYes then
        begin
        ParamOK:=False;
        if (PControl<>nil) then PControl.ActivePage:=Pcontrol.Pages[0];
        Exit ;
        end ;
    end;
  end;
END;

if FileExists(FileName.Text) then
  if HshowMessage(TraduireMemoire('1;Export ETAFI;Le fichier '+FileName.Text+' existe déjà, voulez-vous l''ecraser ?;W;YN;Y;Y'),'','')<>mrYes then
    begin
    ParamOK:=False;
    if (PControl<>nil) then PControl.ActivePage:=Pcontrol.Pages[0];
    FileName.SetFocus ;
    Exit ;
    end ;
end ;

function TOf_ExpEtafi.RecupTypeEcriture: string ;
begin
{$IFDEF SPEC350}
result:=' AND E_QUALIFPIECE="N" ' ;
if Simu then result:='AND (E_QUALIFPIECE="N" OR E_QUALIFPIECE="S") ' ;
{$ELSE}
if CQualif.Value='NOR' then result:=' AND E_QUALIFPIECE="N" '
else if CQualif.Value='NSS' then result:=' AND (E_QUALIFPIECE="N" OR E_QUALIFPIECE="S" OR E_QUALIFPIECE="U" ) '
     else if CQualif.Value='PRE' then result:=' AND E_QUALIFPIECE="P" '
          else if CQualif.Value='SSI' then result:=' AND (E_QUALIFPIECE="S" OR E_QUALIFPIECE="U") '
               else result:='' ;
{$ENDIF}
end ;


Function TOF_ExpEtafi.CreeLeSelect(Exo : string;N : integer): string ;
var Sql: string ;
begin
            Sql:='SELECT E_GENERAL,E_ETABLISSEMENT' ;
if F6 then  Sql:=Sql+',E_AUXILIAIRE' ;
            Sql:=Sql+',SUM(E_DEBIT) AS SUMDEBIT,SUM(E_CREDIT) AS SUMCREDIT';
if N=0 then Sql:=Sql+',E_ECRANOUVEAU' ;
//Sql:=Sql+' FROM ECRITURE LEFT JOIN GENERAUX ON G_GENERAL=E_GENERAL' ;
if (not F6) and (not CCompensation.Checked) then Sql:=Sql+' FROM ECRITURE LEFT JOIN GENERAUX ON G_GENERAL=E_GENERAL WHERE G_COLLECTIF<>"X" '
                                            else Sql:=Sql+' FROM ECRITURE WHERE 1=1 ';
//if F6 then Sql:=Sql+' LEFT JOIN TIERS ON T_AUXILIAIRE=E_AUXILIAIRE' ;
               Sql:=Sql+' AND E_EXERCICE="'+Exo+'" '+RecupTypeEcriture ;
if UnEtab then Sql:=Sql+' AND E_ETABLISSEMENT="'+CEtab.Value+'"';
if N=0 then    Sql:=Sql+' AND E_DATECOMPTABLE>="'+UsDateTime(StrToDate(DatDeb.text))+'" AND E_DATECOMPTABLE<="'+UsDateTime(StrToDate(DatFin.text))+'"' ;
if not CHuit.Checked then Sql:=Sql+' AND E_GENERAL NOT LIKE "8%"' ;
{if GenEtab.Checked then Sql:=Sql+' GROUP BY E_GENERAL,E_ETABLISSEMENT'
                   else} Sql:=Sql+' GROUP BY E_GENERAL,E_ETABLISSEMENT' ;
if F6 then  Sql:=Sql+',E_AUXILIAIRE' ;
if N=0 then Sql:=Sql+',E_ECRANOUVEAU ' ;
{if GenEtab.Checked then Sql:=Sql+' ORDER BY E_GENERAL,E_ETABLISSEMENT'
                   else} Sql:=Sql+' ORDER BY E_GENERAL,E_ETABLISSEMENT' ;
if F6 then  Sql:=Sql+',E_AUXILIAIRE' ;
if N=0 then Sql:=Sql+',E_ECRANOUVEAU ' ;
result:=Sql ;
end ;

procedure TOF_ExpEtafi.ChargeTobG ;
Var Q: TQuery ;
begin
Q:=OpenSql('SELECT G_GENERAL, G_NATUREGENE, G_LIBELLE FROM GENERAUX',True) ;
TobG.LoadDetailDb('TobG','','',Q,True) ;
Ferme(Q) ;
end ;

procedure TOF_ExpEtafi.ChargeTobA ;
Var Q: TQuery ;
begin
Q:=OpenSql('SELECT T_AUXILIAIRE, T_LIBELLE FROM TIERS',True) ;
TobA.LoadDetailDb('TobA','','',Q,True) ;
Ferme(Q) ;
end ;

{JP 08/07/05 : FQ 13318 : Regarde si le cumul des mouvements d'un auxiliaire est créditeur
               ou débiteur et mise à jour du champ AUXCREDDEB
             : REMARQUE : si l'auxiliaire AAAA est créditeur sur l'exercice n et débiteur sur n - 1
                          il figurera sur la ligne 4X1XXXXC pour n et 4X1XXXXD pour n - 1 !!!
                          en attendant une infirmation ....
{---------------------------------------------------------------------------------------}
procedure TOF_ExpEtafi.SetCredDeb(aTob : TOB);
{---------------------------------------------------------------------------------------}
var
  n   : Integer;
  p   : Integer;
  k   : Integer;
  Aux : string;
  Gen : string;
  Eta : string;
  Tmp : TOB;
  Mnt : Double;
begin
  Mnt := 0;
  p := 0;
  for n := 0 to aTob.Detail.Count - 1 do begin
    Tmp := aTob.Detail[n];
    if (Gen <> Tmp.GetString('E_GENERAL')) or (Aux <> Tmp.GetString('E_AUXILIAIRE')) or
       (Eta <> Tmp.GetString('E_ETABLISSEMENT')) then begin
      for k := p to n - 1 do
        if Mnt <= 0 then aTob.Detail[k].AddChampSupValeur('AUXCREDDEB', 'C')
                    else aTob.Detail[k].AddChampSupValeur('AUXCREDDEB', 'D');

      Gen := Tmp.GetString('E_GENERAL');
      Aux := Tmp.GetString('E_AUXILIAIRE');
      Eta := Tmp.GetString('E_ETABLISSEMENT');
      Mnt := 0;
      p := n;
    end;

    Mnt := Arrondi(Mnt + Tmp.GetDouble('SUMDEBIT') - Tmp.GetDouble('SUMCREDIT'), DecLocal);
  end;

  for k := p to aTob.Detail.Count - 1 do
    if Mnt <= 0 then aTob.Detail[k].AddChampSupValeur('AUXCREDDEB', 'C')
                else aTob.Detail[k].AddChampSupValeur('AUXCREDDEB', 'D');
end;

{Récupération des soldes des comptes collectifs}
{---------------------------------------------------------------------------------------}
procedure TOF_ExpEtafi.ChargeTobC (Exo : String ; N : Integer) ;
{---------------------------------------------------------------------------------------}
var
  Sql,
  EcrANO1   : string;
  EcrANO2   : string;
  CreditD   : Double;
  DebitD    : Double;
  CreditC   : Double;
  DebitC    : Double;
  CreditTmp : Double;
  DebitTmp  : Double;
  i         : Integer;
  tTmp      : TOB;
  Q         : TQuery;
  OldEtab   : string;
  NewEtab   : string;
  OldGene   : string;
  NewGene   : string;
  OldAuxi   : string;
  NewAuxi   : string;
  OldCD     : string;

    {On cumul le débit et crédit du tiers en cours en fonction de :
     DebitC et CreditC contiennent le cumuls des debits et crédits des tiers créditeurs sur la période traitée
     DebitD et CreditD contiennent le cumuls des debits et crédits des tiers débiteurs sur la période traitée
    {----------------------------------------------------------}
    procedure AddDebitCredit;
    {----------------------------------------------------------}
    begin
      if (OldCD = 'C') then begin
        DebitC  := DebitC  + DebitTmp;
        CreditC := CreditC + CreditTmp;
      end else begin
        DebitD  := DebitD  + DebitTmp;
        CreditD := CreditD + CreditTmp;
      end;
    end;
begin
  {JP 08/07/05 : FQ 13318 : il faut rajouter l'auxiliaire dans le select, afin de savoir lors du traitement
                 si l'auxiliaire est créditeur ou non}
  Sql := 'SELECT E_GENERAL, E_ETABLISSEMENT, E_AUXILIAIRE, "" AS AUXCREDDEB' +
         ' ,SUM(E_DEBIT) SUMDEBIT, SUM(E_CREDIT) SUMCREDIT';
  {On ne traite les A - Nouveaux que sur l'exercice de référence et pas sur ceux de n - 1 à n - 3}
  if N = 0 then Sql := Sql + ' ,E_ECRANOUVEAU';

  Sql := Sql + ' FROM ECRITURE ' +
               ' LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' +
               {On ne traite icic que l'exercice passé en paramètre et les comptes collectifs}
               ' WHERE E_EXERCICE = "' + Exo + '" AND G_COLLECTIF = "X" ' + RecupTypeEcriture;

  if UnEtab then Sql := Sql + ' AND E_ETABLISSEMENT = "' + CEtab.Value + '"';
  {On peut ne pas traiter l'intégralité de l'exercice de référence}
  if N = 0  then Sql := Sql + ' AND E_DATECOMPTABLE >= "'+ UsDateTime(StrToDate(DatDeb.text)) +
                             '" AND E_DATECOMPTABLE <= "'+ UsDateTime(StrToDate(DatFin.text)) + '"';

  {Les comptes de calsse 8 sont en option}
  if not CHuit.Checked then Sql := Sql + ' AND E_GENERAL NOT LIKE "8%"';

  {Clause GROUP BY}
  Sql := Sql + ' GROUP BY E_GENERAL, E_ETABLISSEMENT, E_AUXILIAIRE';
  if N = 0 then Sql := Sql + ' ,E_ECRANOUVEAU';

  {Clause ORDER BY}
  Sql := Sql + ' ORDER BY E_GENERAL, E_ETABLISSEMENT, E_AUXILIAIRE ' ;
  if N = 0 then Sql := Sql + ', E_ECRANOUVEAU ';

  Q := OpenSql(Sql,True) ;
  TobC.LoadDetailDb('TobC','','',Q,True) ;
  Ferme(Q) ;

  {Mise à jour le champ AUXCREDDEB de la tob en fonction du solde des mouvements sur la
   période de référence et ce par auxiliaire}
  SetCredDeb(TobC);

  DebitTmp  := 0;
  CreditTmp := 0;

  for i := 0 to TobC.Detail.Count - 1 do begin
    tTmp := TobC.Detail[i];
    NewGene := tTmp.GetValue('E_GENERAL');
    NewAuxi := tTmp.GetValue('E_AUXILIAIRE');
    if GenEtab.checked then NewEtab := tTmp.GetValue('E_ETABLISSEMENT');
    if N = 0 then EcrANO1 := tTmp.GetValue('E_ECRANOUVEAU');

    if (NewGene <> OldGene) or (NewEtab <> OldEtab) or (EcrANO1 <> EcrANO2) then begin
      {Cumul du débit et du crédit en fonction du solde général du tiers courant}
      AddDebitCredit;
      {Création de la ligne GénéralC, qui contient les débits et les crédits des auxiliaires créditeurs}
      CreeLigneTobCompensee(OldGene, OldEtab, EcrANO2, 'C', DebitC, CreditC);
      {Création de la ligne GénéralD, qui contient les débits et les crédits des auxiliaires débiteurs}
      CreeLigneTobCompensee(OldGene, OldEtab, EcrANO2, 'D', DebitD, CreditD);

      DebitTmp  := 0;
      CreditTmp := 0;
      DebitC    := 0;
      CreditC   := 0;
      DebitD    := 0;
      CreditD   := 0;
      OldGene   := tTmp.GetValue('E_GENERAL');
      OldAuxi   := tTmp.GetValue('E_AUXILIAIRE');
      if GenEtab.checked then OldEtab   := tTmp.GetValue('E_ETABLISSEMENT');
      if N = 0 then EcrANO2 := tTmp.GetValue('E_ECRANOUVEAU');
      {if N = 0 then 18/07/05} OldCD   := tTmp.GetValue('AUXCREDDEB');
    end
    {On change d'auxiliaire ...}
    else if (NewAuxi <> OldAuxi) then begin
      {Cumul du débit et du crédit en fonction du solde général du tiers précédent}
      AddDebitCredit;
      {Réinitialisation des variables}
      DebitTmp  := 0;
      CreditTmp := 0;
      OldAuxi   := tTmp.GetValue('E_AUXILIAIRE');
      OldCD     := tTmp.GetValue('AUXCREDDEB');
    end;

    {Cumuls des montants}
    DebitTmp  := Arrondi(DebitTmp  + tTmp.GetValue('SUMDEBIT') , DecLocal);
    CreditTmp := Arrondi(CreditTmp + tTmp.GetValue('SUMCREDIT'), DecLocal);
  end;

  {Pour le dernier enregistrement :
   Cumul du débit et du crédit en fonction du solde général du tiers courant}
  AddDebitCredit;
  {Création de la ligne GénéralC, qui contient les débits et les crédits des auxiliaires créditeurs}
  CreeLigneTobCompensee(OldGene, OldEtab, EcrANO2, 'C', DebitC, CreditC);
  {Création de la ligne GénéralD, qui contient les débits et les crédits des auxiliaires débiteurs}
  CreeLigneTobCompensee(OldGene, OldEtab, EcrANO2, 'D', DebitD, CreditD);
end;

procedure TOF_ExpEtafi.CreeLigneTobCompensee(Gene, Etab, EAN, Prefix : string; Deb, Cre:Double);
Var
  T1: Tob;
begin
  if (deb = 0) and (Cre = 0) then Exit;
// Ajoute à TOBC2 les enregistrements
  T1:=CreeLigne ;
  T1.PutValue('Libelle',RecupLib(Gene));
  T1.PutValue('Debit',Deb) ;
  T1.PutValue('Credit',Cre) ;
  T1.PutValue('EcrANouveau',EAN) ;
  T1.PutValue('General',FormateLeGeneral(Gene, Etab, Prefix));
  T1.PutValue('GENERAL2',Gene);
end ;


Function TOF_ExpEtafi.FormateLeGeneral(Gene,Etab,Sens : String ) : String ;
var R1:String ;
begin
  R1:=Gene ;
  if (length(gene)<12)
    then R1:=copy(gene,1,Length(gene)-1)+Sens
    else R1:=copy(gene,1,11)+Sens ;
  Result:=R1 ;
end ;


function TOF_ExpEtafi.CreeLigne : TOB ;
var T1 : TOB ;
begin
T1:=Tob.Create('TobC2',TobC2,-1) ;
T1.AddChampSup('GENERAL', False) ; T1.PutValue('GENERAL', '') ;
T1.AddChampSup('GENERAL2', False) ; T1.PutValue('GENERAL2', '');
T1.AddChampSup('LIBELLE', False) ; T1.PutValue('Libelle','') ;
T1.AddChampSup('DEBIT', False) ; T1.PutValue('DEBIT', arrondi(0,DecLocal));
T1.AddChampSup('CREDIT',False) ; T1.PutValue('CREDIT',arrondi(0,DecLocal));
T1.AddChampSup('ECRANOUVEAU', False) ; T1.PutValue('ECRANOUVEAU', '');
Result:=T1 ;
end;

procedure TOF_ExpEtafi.RecupSolde(Exo : string;N : integer) ;
var QQ : TQuery ;
    Debit,Credit : double ;
    ExoLocal,General,Libelle,EcrAn : string ;
//    TObOK: Tob ;
    i: integer ;
    Gene : String ;
    Aux : String ;
begin
{Si N = 0, on récupère l'exercice sélectionné, sinon on cherche les éventuels exercices de n - 1 à n - 3}
if N <> 0 then ExoLocal := CalculExo(Exo, N)
          else ExoLocal := Exo ;
if ExoLocal='' then Exit ;

//raz Tob
if TobB<>Nil then begin TobB.Free ; TobB:=Nil ; TobB:=Tob.Create('TobB', Nil, -1) ; end else TobB:=Tob.Create('TobB', Nil, -1) ;
If TobA<>nil then begin TobA.Free ; TobA:=Nil ; TobA:=Tob.Create('TobA', Nil, -1) ; end else if F6 then TobA:=Tob.Create('TobA', Nil, -1) ;
If TobG<>nil then begin TobG.Free ; TobG:=Nil ; TobG:=Tob.Create('TobG', Nil, -1) ; end else TobG:=Tob.Create('TobG', Nil, -1) ;
if TobC<>nil then begin TobC.Free ; TobC:=Nil ; TobC:=Tob.Create('TobC', Nil,-1) ; end else if (not CCompensation.Checked) then TobC:=Tob.Create('TobC', Nil,-1) ;
if TobC2<>nil then begin TobC2.Free ; TobC2:=Nil ; TobC2:=Tob.Create('TobC2', Nil,-1) ; end else if (not CCompensation.Checked) then TobC2:=Tob.Create('TobC2', Nil,-1) ;
QQ:=OpenSql(CreeLeSelect(ExoLocal,N),True) ;
TobB.LoadDetailDB('Ecr','','',qq,true) ;
ferme (QQ) ;

F6:=CFormat.ItemIndex=1 ;
ChargetobG ; // Les libelles des généraux
if F6 then ChargeTobA ; // les libellés des auxiliaires
if (not F6) and  (Not CCompensation.Checked) then ChargeTobC(ExoLocal,N) ; // si compensation

try
  if (CFormat.ItemIndex=0) and (not CCompensation.Checked) then F6:=FALSE ;
  for i:=0 to TobB.Detail.Count-1 do
    begin
    // Remplit les valeurs nécessaires
    Aux:='' ;

    if F6 then if VarType(TobB.detail[i].GetValue('E_AUXILIAIRE'))=VarNull then Aux:='' else Aux:=TobB.detail[i].GetValue('E_AUXILIAIRE') ;
    Gene:=TobB.detail[i].GetValue('E_general') ;
    General:=Gene ;

    if N<>0 then EcrAn:='' else EcrAn:=TobB.Detail[i].GetValue('E_ECRANOUVEAU') ;
    Debit  :=arrondi(TobB.Detail[i].GetValue('SUMDEBIT'),DecLocal) ;
    Credit :=arrondi(TobB.Detail[i].GetValue('SUMCREDIT'),DecLocal) ;
    if (Gene=TBenef.Text) or (Gene=TPerte.Text) then begin Gene:=TResult.Text ; Libelle:=RecupLib(Gene) ; end
      else if Aux<>'' then Libelle:=RecupLibaux(Aux) else Libelle:=RecupLib(gene) ;
      RemplitTob(General,Aux,libelle,EcrAn,N,Debit,Credit) ;
    end ;

  if (not F6) and  (Not CCompensation.Checked) then
    for i:=0 to TobC2.Detail.count-1 do
    begin
    General:=TobC2.detail[i].GetValue('General') ;
    Libelle:=TobC2.detail[i].GetValue('Libelle') ;
    Aux:='' ;
    Debit  :=arrondi(TobC2.Detail[i].GetValue('DEBIT'),DecLocal) ;
    Credit :=arrondi(TobC2.Detail[i].GetValue('CREDIT'),DecLocal) ;
    if N<>0 then EcrAn:='' else EcrAn:=TobC2.Detail[i].GetValue('ECRANOUVEAU') ;
//    TobOK :=
    RemplitTob(General,Aux,libelle,EcrAn,N,Debit,Credit);
    {JP 18/07/05 : cette rustine fonctionnait mal et n'est plus nécessaire car les modifications
                   nécéssaires ont été apportées dans ChargeTobC pour gérer la non compensation
                   des collectifs clients et fournisseurs
                   : REMARQUE : si l'auxiliaire AAAA est créditeur sur l'exercice n et débiteur sur n - 1
                                il figurera sur la ligne 4X1XXXXC pour n et 4X1XXXXD pour n - 1 !!!
                                en attendant une infirmation ....
     RetraiteSolde(TobOK, General, Exo); // FQ 13318}
    end ;
finally
end ;
end ;

procedure TOF_ExpEtafi.GenEtabOnClick (Sender : TobJect) ;
begin
  CPTAGEN.Enabled := genEtab.Checked;
end ;

function TOF_ExpEtafi.RemplitTob(General,Aux,Libelle,EcrAn: string;N: integer; Debit,Credit: double) : Tob;
var T1: Tob ;
begin
T1:=nil; Result := nil;
if F6 and (Aux<>'') then T1:=TheTob.FindFirst(['GENERAL','AUXILIAIRE'],[General,Aux],True) ;
if F6 and (Aux='')  then T1:=TheTob.FindFirst(['AUXILIAIRE'],[General],True) ;
if not F6           then T1:=TheTob.FindFirst(['GENERAL'],[General],True) ;
if T1=nil then  T1:=CreeLaFille ;
if T1=nil then Exit ;
T1.PutValue('SGENERAL',General) ;
T1.PutValue('LIBELLE',Libelle) ;
if not F6 or (Aux<>'') then T1.PutValue('GENERAL',General) else T1.PutValue('AUXILIAIRE',General) ;
if Aux<>'' then T1.PutValue('AUXILIAIRE',Aux) ;
case N of
  0 :  begin
       if (EcrAn='OAN') OR (EcrAn='H') then begin
         T1.PutValue('SOLDEAN',arrondi(T1.GetValue('SOLDEAN')+Debit-Credit,DecLocal)) ;
         end
       else begin
         T1.PutValue('MVDEBIT',arrondi(T1.GetValue('MVDEBIT')+Debit,DecLocal)) ;
         T1.PutValue('MVCREDIT',arrondi(T1.GetValue('MVCREDIT')+Credit,DecLocal)) ;
       end ;
       T1.PutValue('SOLDE', arrondi(T1.GetValue('SOLDEAN')+T1.GetValue('MVDEBIT')-T1.GetValue('MVCREDIT'),DecLocal));
       end ;
  1 :  T1.PutValue('SOLDEN1',arrondi(T1.GetValue('SOLDEN1')+Debit-Credit,DecLocal)) ;
  2 :  T1.PutValue('SOLDEN2',arrondi(T1.GetValue('SOLDEN2')+Debit-Credit,DecLocal)) ;
  3 :  T1.PutValue('SOLDEN3',arrondi(T1.GetValue('SOLDEN3')+Debit-Credit,DecLocal)) ;
  end ;

Result := T1;
end ;

procedure TOF_ExpEtafi.ExportLaTob(FName : string) ;
var FDest : TextFile ;Taille,i : integer ; Ligne : string ;
begin
AssignFile(FDest, FName ) ;
ReWrite(FDest) ;
TheTob.Detail.Sort('SGENERAL') ;
if F6 then Taille:=40 else Taille:=35;
for i:=0 to TheTob.Detail.Count-1 do
  begin
  Ligne:='';
  if F6 then Ligne:=Ligne+format('%-16.16s',[TheTob.detail[i].GetValue('AUXILIAIRE')]) ;
  Ligne:=Ligne+format('%-12.12s',[TheTob.detail[i].GetValue('GENERAL')]) ;
  if F6 then Ligne:=Ligne+'    ' ; //racine auxiliaire 4 car.
  Ligne:=Ligne+format('%-*.*s',[Taille,Taille,TheTob.detail[i].GetValue('LIBELLE')]) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('SOLDEAN') ),RSeparateur.ItemIndex,DecLocal) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('MVDEBIT') ),RSeparateur.ItemIndex,DecLocal) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('MVCREDIT')),RSeparateur.ItemIndex,DecLocal) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('SOLDE')   ),RSeparateur.ItemIndex,DecLocal) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('SOLDEN1') ),RSeparateur.ItemIndex,DecLocal) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('SOLDEN2') ),RSeparateur.ItemIndex,DecLocal) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('SOLDEN3') ),RSeparateur.ItemIndex,DecLocal) ;
  WriteLn(FDest,Ligne) ;(*Ansi2Ascii(Ligne)) ; *)
  end ;
CloseFile(FDest) ;
end;

procedure TOF_ExpEtafi.OnChangeCFormat(Sender: TObject) ;
begin
EnableButtons ;
end ;

procedure TOF_ExpEtafi.OnChangeCExo(Sender: TObject) ;
begin
if (DatDeb=nil) or (DatFin=nil) then Exit ;
ExoToDates(CExo.value,DatDeb,DatFin) ;
end ;

{Recherche des (éventuels) éxercices de n à n - 3 par rapport à l'exercice sélectionné
{---------------------------------------------------------------------------------------}
function TOF_ExpEtafi.CalculExo(Exo : string; N : integer) : string;
{---------------------------------------------------------------------------------------}
var
  i : integer ;
  TabExo : array[1..8] of string ;
begin
  Result := '';
  {1. Remplissage de TABEXO avec les 8 exercices les plus récents dans l'ordre décroissant,
      s'ils existent bien sûr}
  TabExo[1] := VH^.Suivant  .Code; {Dernier exercice ouvert}
  TabExo[2] := VH^.EnCours  .Code; {Premier exercice ouvert ou exercice clôturé provisoirement}
  TabExo[3] := VH^.Precedent.Code; {Dernier exercice clôturé définitivement}
  {VH^.Exoclo est un tableau de 5 qui contient les exercices clôturés les plus récents sauf
   s'il y a plus de 5 exercices clôturés, à ce moment le tableau contient les exercices clôturés
   de 2 à 6, l'exercice clôturé le plus récent étant déjà dans VH^.Précédent}
  for i := 1 to 5 do
    if VH^.Exoclo[i].Code <> VH^.Precedent.Code then
      TabExo[i+3] := VH^.Exoclo[i].Code;

  {2. On renvoie les exercices de n à n - 3 selon le paramètre "N"
      qui a une valeur comprise entre 0 et 3}
  for i := 1 to 8 - N do 
    {Si TabExo[i] correspond à l'exercice sélectionné dans la fiche ...}
    if TabExo[i] = Exo then
      {On renvoie l'exercice n - "N"}
      Result := TabExo[i + N];
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_ExpEtafi.GereEtablissement;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(CEtab) then begin
    {Si l'on ne gère pas les établissement ...}
    if not VH^.EtablisCpta  then begin
      {... on affiche l'établissement par défaut}
      CEtab.Value := VH^.EtablisDefaut;
      {... on désactive la zone}
      CEtab.Enabled := False;
    end

    {On gère l'établisement, donc ...}
    else begin
      {... On commence par regarder les restrictions utilisateur}
      PositionneEtabUser(CEtab);
      {... s'il n'y a pas de restrictions, on reprend le paramSoc
       JP 25/10/07 : FQ 19970 : Finalement on oublie l'option de l'établissement par défaut
      if CEtab.Value = '' then begin
        {... on affiche l'établissement par défaut
        CEtab.Value := VH^.EtablisDefaut;
        {... on active la zone
        CEtab.Enabled := True;
      end;}
    end;
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_ExpEtafi.ControlEtab;
{---------------------------------------------------------------------------------------}
var
  Eta : string;
begin
  if not Assigned(CEtab) then Exit;
  {S'il n'y a pas de gestion des établissement, logiquement, on ne force pas l'établissement !!!}
  if not VH^.EtablisCpta then Exit;
  
  Eta := EtabForce;
  {S'il y a une restriction utilisateur et qu'elle ne correspond pas au contenu de la combo ...}
  if (Eta <> '') and (Eta <> CEtab.Value) then begin
    {... on affiche l'établissement des restrictions}
    CEtab.Value := Eta;
    {... on désactive la zone}
    CEtab.Enabled := False;
  end;
end;

(* JP 18/07/05 : a priori ce n'est plus nécessaire suite aux modifications apportées à ChargeTobC
  // FQ 13318
procedure TOF_ExpEtafi.RetraiteSolde(TobOK: Tob; const General, Exo: String);
var
  Q : TQuery;
  szSQL, szGen, szSens : String;
  dblSum, dblTotal1, dblTotal2 : Double;
begin
  //RR 02/06/2005 - on laisse telquel
  //FQ 15569
  //szGen := Copy(General, 1, Length(General)-1) + '0' ; //GetParamSocSecur('SO_BOURREGEN','');
  szGen := General ;
  szSens := General[Length(General)];
  if TobOK.GetNumChamp('GENERAL2') > 0 then szGen := TobOK.GetString('GENERAL2');

  szSQL := 'SELECT SUM(E_DEBIT)-SUM(E_CREDIT) SOMME FROM ECRITURE '+
           'WHERE E_EXERCICE="'+Exo+'" AND E_GENERAL="'+szGen+'" '+
           'AND E_DATECOMPTABLE>="'+UsDateTime(StrToDate(DatDeb.text))+'" AND E_DATECOMPTABLE<="'+UsDateTime(StrToDate(DatFin.text))+'" '+
           'GROUP BY E_AUXILIAIRE';

  Q := OpenSql(szSQL, True);
  dblTotal1 := 0;
  dblTotal2 := 0;
  while not Q.Eof do begin
    dblSum := Q.FindField('SOMME').AsFloat;
    if (dblSum < 0) then dblTotal1 := dblTotal1+dblSum
                    else dblTotal2 := dblTotal2+dblSum;
    Q.Next;
  end;
  if (szSens = 'C') then TobOK.PutValue('SOLDE', dblTotal1)
                    else TobOK.PutValue('SOLDE', dblTotal2);
  Ferme(Q);
end;
*)

Initialization
registerclasses([TOF_ExpEtafi]);
end.
