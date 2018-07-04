unit TofMulRupRub;

interface
uses  Controls,StdCtrls,Graphics,Classes,sysutils,dbTables,comctrls,
      HCtrls,ent1,HEnt1,HMsgBox,UTOF,LookUp,HDB,FE_Main,Vierge,
      HRichOLE,YpUtils,RubFam ;

type TRupture = class
     Nature,Plan,Classe,Libelle : string ;
     end ;

type TOF_MulRupRub = class (TOF)
     private
       XX_Where : THEdit ;
       NatureRupt,PlanRupt,LaFamille : THValComboBox ;
       FListe : THDBGrid ;
       Tous  : boolean ;
       Pages: TPageControl;
       procedure OnNatureRuptChange(Sender : TObject) ;
       procedure OnPlanRuptChange(Sender : TObject) ;
       procedure ChargePlanRupt ;
       procedure MiseAJourDeXXWhere ;
       procedure OnBOuvrirClick(Sender : TObject) ;
       procedure OnBZoomClick(Sender : TObject) ;
       procedure TraiteRubriqueGeneOuAuxi(Rupture : TRupture) ;
       procedure TraiteRubriqueAnalytique(Rupture : TRupture) ;
       procedure InsertOrUpdateRubriques(Nom,Libelle,Famille,TypeRub,Compte1,Axe : string ) ;
       function AutoriseMaj(Nom : String) : boolean ;
       function RecupAxe(Nature : string ;var Axe : string;var LonAxe : integer;var NumAxe :  TFichierBase ) : boolean;
       procedure RecupSousSection(NumAxe : TFichierBase;Plan : string ; var PlanDeb,PlanLon : integer ) ;
       procedure OnElipsisClickLafamille(Sender : TObject) ;
     public
       procedure OnLoad ; override ;
       procedure OnClose ; override ;
     end;

const NATUREGEN  = 'RUG'  ;
      NATUREAUX  = 'RUT'  ;
      TYPERUBGE  = 'GEN'  ;
      TYPERUBAN  = 'ANA'  ;
      MODESOLDE  = '(SM)' ;
      SIGNERUBR  = 'POS'  ;
      NATURRUBR  = 'CPT'  ;

      TMsg : array[1..4] of string = (
      {01}        '1;Conversion des ruptures en rubriques;La rubrique %% existe déjà voulez-vous la remplacer;Q;YNL;Y;Y'
      {02}        ,'2;Conversion des ruptures en rubriques;Vous devez renseigner la famille de rubriques;W;O;O;O'
      {03}        ,'2;Conversion des ruptures en rubriques;Vous devez sélectionner au moins une rupture  ;W;O;O;O'
      {04}        ,'2;Conversion des ruptures en rubriques;La conversion des ruptures a été effectuée;A;O;O;O'
                  );
                  
implementation

Uses FamRub ;

procedure TOF_MulRupRub.OnLoad;
var BButton : TButton ;
begin
XX_Where:=THedit(GetControl('XX_WHERE')) ;
NatureRupt:=THValComboBox(GetControl('NATURERUPT')) ;
if (NatureRupt<>nil) and not assigned(NatureRupt.OnChange) then NatureRupt.OnChange:=OnNatureRuptChange ;
PlanRupt:=THValComboBox(GetControl('PLANRUPT')) ;
if (PlanRupt<>nil) and not assigned(PlanRupt.OnChange) then PlanRupt.OnChange:=OnPlanRuptChange ;
FListe:=THDBGrid(GetControl('FLISTE')) ;
BButton:=TButton(GetControl('BOUVRIR')) ;
if (BButton<>nil) (*and not assigned(BButton.OnClick)*) then BButton.OnClick:=OnBOuvrirClick ;
BButton:=TButton(GetControl('BZOOM')) ;
if (BButton<>nil) and not assigned(BButton.OnClick) then BButton.OnClick:=OnBZoomClick ;
LaFamille:=THValComboBox(GetControl('LAFAMILLE')) ;
if (LaFamille<>nil) and not assigned(LaFamille.OnElipsisClick) then LaFamille.OnElipsisClick:=OnElipsisClickLafamille ;
Pages:=TPageControl(GetControl('PAGES')) ;
end;

procedure TOF_MulRupRub.OnElipsisClickLafamille(Sender : TObject) ;
Var Fam,LibFam : String ;
begin
(*
LaFAmille.Text:=ChoisirMulti('Famille de rubriques','CHOIXCOD','CC_CODE','CC_LIBELLE','CC_TYPE="RBF"','','') ;
if LaFAmille.Text='???' then LaFAmille.Text:='' ;
*)
Fam:=LaFAmille.Text ;
ParametrageFamilleRubrique('','',Fam,LibFam,False,TRUE) ;
If Fam<>'' Then LaFAmille.Text:=Fam ;

end ;



procedure TOF_MulRupRub.OnClose ;
Begin
end;

procedure TOF_MulRupRub.ChargePlanRupt ;
var Q : TQuery ;
begin
Q:=OpenSql('SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+NatureRupt.Value+'"',True) ;
PlanRupt.Clear ;
While not Q.Eof do
  begin
  PlanRupt.Items.Add(Q.Findfield('CC_LIBELLE').AsString) ;
  PlanRupt.Values.Add(Q.Findfield('CC_CODE').AsString) ;
  Q.next ;
  end ;
ferme(Q) ;
end ;

procedure TOF_MulRupRub.OnPlanRuptChange(Sender : TObject) ;
begin
MiseAJourDeXXWhere ;
end ;

procedure TOF_MulRupRub.OnNatureRuptChange(Sender : TObject) ;
begin
ChargePlanRupt ;
MiseAJourDeXXWhere ;
end ;

procedure TOF_MulRupRub.MiseAJourDeXXWhere ;
begin
XX_Where.Text:='' ;
if NatureRupt.Value<>'' then
  XX_Where.Text:=XX_Where.Text+'RU_NATURERUPT="'+NatureRupt.Value+'" ';
if PlanRupt.Value<>'' then
  begin
  if NatureRupt.Value<>'' then XX_Where.Text:=XX_Where.Text+'AND ' ;
  XX_Where.Text:=XX_Where.Text+'RU_PLANRUPT="'+PlanRupt.Value+'" ' ;
  end ;
end ;

procedure TOF_MulRupRub.OnBZoomClick(Sender : TObject) ;
begin
FamillesRub ;
end ;

procedure TOF_MulRupRub.OnBouvrirClick(Sender : TObject) ;
var i : integer ; Q : TQuery ; Rupture : TRupture ;
begin
if FListe.nbSelected<1 then begin HShowMessage(TraduireMemoire(TMsg[3]),'','') ; Exit ; end ;
Q:=TQuery(FListe.DataSource.DataSet) ;
Tous:=False ;
Rupture:=TRupture.Create ;
while FListe.nbSelected>0 do
  begin
  FListe.GotoLeBookmark(0) ;
  Rupture.Nature:=Q.Findfield('RU_NATURERUPT').AsString ;
  Rupture.Plan  :=Q.Findfield('RU_PLANRUPT').AsString ;
  Rupture.Classe:=Q.Findfield('RU_CLASSE').AsString ;
  Rupture.Libelle:=Q.Findfield('RU_LIBELLECLASSE').AsString ;
  if (Rupture.Nature=NATUREGEN) or (Rupture.Nature=NATUREAUX) then
    begin
    if LaFamille.Value='' then begin HShowMessage(TraduireMemoire(TMsg[2]),'','') ; Pages.ActivePage:=Pages.Pages[4] ; LaFamille.SetFocus ; break ; end
                          else TraiteRubriqueGeneOuAuxi(Rupture)
    end
    else TraiteRubriqueAnalytique(Rupture) ;
  FListe.FlipSelection ;
  end ;
Rupture.Free ;
if FListe.nbSelected<=0 then HShowMessage(TraduireMemoire(TMsg[4]),'','') ;
end ;

procedure TOF_MulRupRub.TraiteRubriqueGeneOuAuxi(Rupture : TRupture) ;
var Nom,Libelle,Famille,TypeRub,Compte1,Axe : string ;
begin
Nom:=Rupture.Nature+Rupture.Plan+Rupture.Classe ;
Libelle:=Rupture.Libelle ;
Famille:=LaFamille.Value  ;
TypeRub:= TYPERUBGE;
Axe:='' ;
Compte1:=copy(Rupture.Classe,1,length(Rupture.Classe)-1)+MODESOLDE ;
InsertOrUpdateRubriques(Nom,Libelle,Famille,TypeRub,Compte1,Axe) ;
end ;

procedure TOF_MulRupRub.TraiteRubriqueAnalytique(Rupture : TRupture) ;
var Tmp,Nom,Libelle,Famille,TypeRub,Compte1,Axe,Liste,Plan : string ; Q : TQuery ; i,PlanLon,PlanDeb,LonAxe : integer ;NumAxe : TFichierBase ;
begin
if not RecupAxe(Rupture.Nature,Axe,LonAxe,NumAxe) then Exit ;
Nom:=Rupture.Nature+Rupture.Plan+Rupture.Classe ;
Libelle:=Rupture.Libelle ;
Famille:='CBS'  ;
TypeRub:= TYPERUBAN;
Q:=OpenSql('SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+Rupture.Nature+'" AND CC_CODE="'+Rupture.Plan+'"',True) ;
if Not Q.Eof then Liste:=Q.findfield('CC_LIBRE').AsString else Liste:='' ;
ferme(Q) ;
if Liste='' then Compte1:=copy(Rupture.Classe,1,length(Rupture.Classe)-1)
else
  begin
  memset(Compte1,'?',LonAxe) ;
  Plan:=ReadTokenSt(Liste) ;
  while Plan<>'' do
    begin
    RecupSousSection(NumAxe,Plan,PlanDeb,PlanLon) ;
    Tmp:=Rupture.Classe ;
    for i:=1 to PlanLon do begin Compte1[PlanDeb-1+i]:=Tmp[i] end ;
    Plan:=ReadTokenSt(Liste) ;
    end ;
  for i:=LonAxe downto 0 do if Compte1[i]='?' then Compte1[i]:=' ' else break ;
  Compte1:=trim(Compte1) ;
  end ;
if Compte1<>'' then
  begin
  Compte1:=Compte1+MODESOLDE ;
  InsertOrUpdateRubriques(Nom,Libelle,Famille,TypeRub,Compte1,Axe) ;
  end ;
end ;


procedure TOF_MulRupRub.InsertOrUpdateRubriques(Nom,Libelle,Famille,TypeRub,Compte1,Axe : string ) ;
var Q : TQuery ; Okok  : boolean ;
begin
Okok:=False ;
Nom:='RU:'+copy(Nom,1,14) ;
if (Nom<>'') then
  begin
  Q:=OpenSql('SELECT * FROM RUBRIQUE WHERE RB_RUBRIQUE="'+Nom+'"',False) ;
  if Q.Eof then begin Okok:=True ; Q.Insert ; InitNew(Q) ; Q.FindField('RB_RUBRIQUE').AsString  :=Nom ; end
           else begin Okok:=AutoriseMaj(Nom) ; Q.Edit ; end ;
  if Okok then
    begin
    Q.FindField('RB_LIBELLE').AsString   :=Libelle ;
    Q.FindField('RB_FAMILLES').AsString  :=Famille ;
    Q.FindField('RB_SIGNERUB').AsString  :=SIGNERUBR ;
    Q.FindField('RB_TYPERUB').AsString   :=TypeRub ;
    Q.FindField('RB_COMPTE1').AsString   :=Compte1 ;
    Q.FindField('RB_AXE').AsString       :=Axe ;
    Q.FindField('RB_NATRUB').AsString    :=NATURRUBR ;
    Q.FindField('RB_CODEABREGE').AsString:=copy(Nom,4,6) ;
    // JLD + GP Pour XX_PREDEFINI
    Q.FindField('RB_PREDEFINI').AsString:='DOS' ; Q.FindField('RB_NODOSSIER').AsString:='000000' ;
    if V_PGI_Env<>Nil then if V_PGI_Env.ModeFonc<>'MONO' then Q.FindField('RB_NODOSSIER').AsString:=V_PGI_ENV.NoDossier ;
    Q.Post ;
    end ;
  ferme(Q) ;
  end ;
end ;

function TOF_MulRupRub.AutoriseMaj(Nom : String) : boolean ;
var Ret : Word ;
begin
if not Tous then
  begin
  Ret:=HShowMessage(TraduireMemoire(TMsg[1]),Nom,'') ;
  if Ret=mrAll then Tous:=True ;
  if (Ret=mrYes) or (Ret=mrAll) then Result:=True else Result:=False ;
  end
  else
    Result:=True ;
end ;

function TOF_MulRupRub.RecupAxe(Nature : string ;var Axe : string;var LonAxe : integer;var NumAxe :  TFichierBase ) : boolean;
begin
Result:=True ;
if Nature='RU1' then begin Axe:='A1' ; NumAxe:=fbAxe1 ; end
else if Nature='RU2' then begin Axe:='A2' ; NumAxe:=fbAxe2 ; end
     else if Nature='RU3' then begin Axe:='A3' ; NumAxe:=fbAxe3 ; end
          else if Nature='RU4' then begin Axe:='A4' ; NumAxe:=fbAxe4 ; end
               else begin Axe:='' ; Result:=False ; end;
if Result then LonAxe:=VH^.Cpta[NumAxe].Lg;
end ;

procedure TOF_MulRupRub.RecupSousSection(NumAxe : TFichierBase;Plan : string ; var PlanDeb,PlanLon : integer ) ;
var i : integer ;
begin
PlanDeb:=0 ; PlanLon:=0 ;
for i:=1 to MaxSousPlan do
  begin
  if VH^.SousPlanAxe[NumAxe,i].Code=Plan then
    begin
    PlanDeb:=VH^.SousPlanAxe[NumAxe,i].Debut ;
    PlanLon:=VH^.SousPlanAxe[NumAxe,i].Longueur ;
    break ;
    end ;
  end ;
end ;

Initialization
registerclasses([TOF_MulRupRub]);
end.
