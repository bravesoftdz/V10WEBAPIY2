unit VerEuro;

interface
uses Classes,UTOB,UTOF,dbTables,Hctrls,Ent1,HMsgBox,Hent1,UtilPgi,SysUtils,FE_main,AffList,paramsoc ; 

type TOF_VerEuro = class (TOF)
     private
       FExo,FEtab,FNatJnl1,FNatJnl2 : THValComboBox;
       SaufJnl,FDateEcr1,FDateEcr2 : THEdit ;
       procedure OnClickElipsis(Sender : TObject) ;
     public
       procedure OnLoad ; override ;
       procedure OnUpdate ; override ;
       procedure OnClose ; override ;
     end;

type VER_EURO = class(TObject)
      private
        ParamOk: boolean ;
        TobJalPer,TobErreur,TobQuotite,TobPiece,TobEccDeb,TobEccCre,TobEcart: TOB ;
        TypeSaisie,EccEuroDebit,EccEuroCredit,Benef,Perte,LaSai: string  ;
        AddFra,AddEur,AddDev: double ;
        procedure AjoutUneLigneDEcartDeConversion(TobEcr:tob;DebitOuCredit:string; EcartFra,EcartEur,EcartDev: double) ;
        function VerificationMontant(var F,E,D:double;Taux: double; General,QuelSaisie: string;Sens: boolean): integer ;
        function VerificationConvertEcriture(TobEcr: TOB;SurPiece: boolean): integer;
        function VerificationConvertPiece(TobEcr: TOB): integer;
        function VerificationConvertJalPer(TobEcr: TOB): integer ;
        procedure VerificationConvertCompta();
        function CheckParam: boolean ;
        function ChangePiece(TobEcr: TOB): boolean ;
        function CheckCompte(Compte: string): integer ;
        procedure AjoutErreur(General: string;Err: integer ;F,E,D,Tmp,Taux,Quotite: double) ;
        function QuelQuotite(Devise: string): double;
        procedure initvaleur ;
      public
        Exo,Etab,Jal1,Jal2,SaufJal: string ;
        DateDeb,DateFin: TDateTime ;
        constructor Create;
        destructor Destroy; override ;
        procedure execute ;
        procedure affiche ;
      END ;

implementation
function VER_EURO.QuelQuotite(Devise: string): double;
var LaTob: Tob ;
begin
result:=1 ;
if TobQuotite.Detail.Count<>0 then begin
  LaTob:=TobQuotite.FindFirst(['D_DEVISE'],[Devise],False) ;
  if LaTob<>nil then Result:=LaTob.GetValue('D_QUOTITE') ;
  end ;
end ;

procedure VER_EURO.AjoutUneLigneDEcartDeConversion(TobEcr:Tob;DebitOuCredit:string;EcartFra,EcartEur,EcartDev: double) ;
var Dev,Eur,Fra: double ; LaTob: Tob;
begin
if DebitOuCredit='DEBIT' then LaTob:=TobEccDeb else LaTob:=TobEccCre ;
Dev:=EcartDev;
if VH^.TenueEuro then begin Fra:=EcartEur; Eur:=EcartFra; end
                 else begin Fra:=EcartFra; Eur:=EcartEur; end ;
if (LaTob<>nil) and (Fra<=0) and (Eur<=0) and (Dev<=0) then begin LaTob.Free ; LaTob:=nil ; end ;
if (Fra>0) or (Eur>0) or (Dev>0) then begin
  if laTob=nil then begin
    if TobEcart=nil then TobEcart:=Tob.create('ECRITURE',nil,-1) ;
    LaTob:=Tob.Create('$UNKNOWN',TobEcart,-1) ;
    with LaTob do begin
      Dupliquer(TobEcr,false,true) ;
      if GetValue('E_ECRANOUVEAU')='N' then begin if DebitOuCredit='DEBIT' then PutValue('E_GENERAL',EccEuroDebit) else PutValue('E_GENERAL',EccEuroCredit) ; end
                                       else begin if DebitOuCredit='DEBIT' then PutValue('E_GENERAL',Perte) else PutValue('E_GENERAL',Benef) ; end ;
         end ;
    end;
  with LaTob do begin
    PutValue('E_'+DebitOuCredit,Fra) ;
    PutValue('E_'+DebitOuCredit+'EURO',Eur) ;
    PutValue('E_'+DebitOuCredit+'DEV',Dev) ;
    end ;
  end ;
end ;

{*
Result :
       0: Tout est OK
      -1: Erreur de convertion entre F  => E (Saisie en Francs)
      -2: Erreur de convertion entre F  => D (Saisie en Francs)
      -3: Erreur de convertion entre E  => F (Saisie en Euro)
      -4: Erreur de convertion entre E  => D (Saisie en Euro)
      -5: Erreur de convertion entre D  => E (Saisie en Devise)
      -5: Erreur de convertion entre E° => F (Saisie en Devise)

      °: Montant convertit sans arrondi (D to E).
*}
function VER_EURO.VerificationMontant(var F,E,D:double;Taux: double; General,QuelSaisie: string;Sens: boolean): integer ;
var Err: integer ; De,Eu,Fr,DevConv,Tmp,Quotite: double ;
begin
Err:=0 ; tmp:=0 ; Quotite:=0 ;
if VH^.TenueEuro then begin Fr:=E; Eu:=F; De:=D; end else begin Fr:=F; Eu:=E ; De:=D ; end ;
//Saisie en Devise
if ((QuelSaisie<>'FRA') and (QuelSaisie<>'EUR')) then begin
  Quotite:=QuelQuotite(QuelSaisie) ;
  Tmp:=DEVISETOEURO(De,Taux,Quotite) ;
  if Tmp<>Eu then begin AjoutErreur(General,-5,Fr,Eu,De,Tmp,Taux,Quotite) ; Err:=-5 ;   Eu:=Tmp ; end ;
  DevConv:=DEVISETOEURONA(De,Taux,Quotite) ;
  Tmp:=EUROTOFRANC(DevConv) ;
  if Tmp<>Fr then begin  AjoutErreur(General,-6,Fr,Eu,De,Tmp,Taux,Quotite) ; Err:=-6 ;  Fr:=Tmp ; end ;
  end ;
// Saisie en Francs
if (QuelSaisie='FRA') then begin
  Tmp:=FRANCTOEURO(Fr) ;
  if Tmp<>Eu then begin   AjoutErreur(General,-1,Fr,Eu,De,Tmp,Taux,Quotite) ; Err:=-1 ; Eu:=Tmp ; end ;
  if Fr<>De  then begin   AjoutErreur(General,-2,Fr,Eu,De,Tmp,Taux,Quotite) ; Err:=-2 ; De:=Fr ;  end ;
  end ;
// Saisie en Euro
if (QuelSaisie='EUR') then begin
  Tmp:=EUROTOFRANC(Eu) ;
  if Tmp<>Fr then begin  AjoutErreur(General,-3,Fr,Eu,De,Tmp,Taux,Quotite) ; Err:=-3 ; Fr:=Tmp ; end ;
  if Eu<>De  then begin  AjoutErreur(General,-4,Fr,Eu,De,Tmp,Taux,Quotite) ; Err:=-4 ; De:=Eu ;  end ;
  end ;
if Err<>0 then AjoutErreur(General+' Après',Err,Fr,Eu,De,0,Taux,Quotite) ;
if VH^.TenueEuro then begin E:=Fr; F:=Eu; D:=De; end else begin F:=Fr; E:=Eu ; D:=De ; end ;
Result:=Err ;
end ;

function VER_EURO.VerificationConvertEcriture(TobEcr: TOB;SurPiece: boolean): integer;
var Err,Quotite: integer; Taux,Fra,Eur,Dev: double ; Lib,DevSaisie,DevDossier: string ;
    Tab: array[1..2] of string ; i: integer ;
begin
Result:=0 ; Err:=0 ; Tab[1]:='DEBIT' ; Tab[2]:='CREDIT' ;
with TobEcr do begin
  if CheckCompte(GetValue('E_GENERAL'))=-1 then begin TobEccDeb:=TobEcr ; Exit ; end
    else if CheckCompte(GetValue('E_GENERAL'))=1 then begin TobEccCre:=TobEcr ; Exit ; end ;
  DevSaisie:=GetValue('E_DEVISE') ;
  if GetValue('E_DEVISE')=V_PGI.DevisePivot then begin
    if (GetValue('E_SAISIEEURO')='X') then begin if (VH^.TenueEuro=True) then DevSaisie:='FRA' else DevSaisie:='EUR' ; end
                                      else begin if (VH^.TenueEuro=True) then DevSaisie:='EUR' else DevSaisie:='FRA' ; end ;
    end;
  // Vérification Débit/Crédit
  for i:=1 to 2 do begin
    Fra:=GetValue('E_'+tab[i]) ;    Eur:=GetValue('E_'+tab[i]+'EURO') ;    Dev:=GetValue('E_'+tab[i]+'DEV') ;
    Taux:=GetValue('E_TAUXDEV') ; Lib:=GetValue('E_GENERAL')+' '+inttostr(GetValue('E_NUMEROPIECE')) ;
    Err:=VerificationMontant(Fra,Eur,Dev,Taux,lib,DevSaisie,False) ;
    if Err<>0 then begin
      Result:=Result+1 ;
      if SurPiece then begin
      AjoutUneLigneDecartDeConversion(tobecr,tab[i],Fra-GetValue('E_'+tab[i]),Eur-GetValue('E_'+tab[i]+'EURO'),Dev-GetValue('E_'+tab[i]+'DEV') ) ;
      end
      else begin
      PutValue('E_'+tab[i],Fra) ;     PutValue('E_'+tab[i]+'EURO',Eur) ;     PutValue('E_'+tab[i]+'DEV',Dev) ;
      end ;
    end ;
  end ;
end ;
end ;

function VER_EURO.VerificationConvertPiece(TobEcr: TOB): integer;
var tmp: double ; Tab: array[1..6] of string ; i: integer;
begin
result:=0 ;
Tab[1]:='E_DEBIT'; Tab[2]:='E_DEBITEURO'; Tab[3]:='E_DEBITDEV' ;
Tab[4]:='E_CREDIT';Tab[5]:='E_CREDITEURO';Tab[6]:='E_CREDITDEV' ;
if TobPiece=nil then begin
  TobPiece:=Tob.Create('$INCONNU',nil,-1) ;
  TobPiece.Dupliquer(TobEcr,False,true) ;
  end ;
if ChangePiece(TobEcr) then begin
  Result:=Result+VerificationConvertEcriture(TobPiece,True) ;
  TobPiece.Dupliquer(TobEcr,False,true) ;
  end ;
for i:=1 to 6 do begin
  Tmp:=TobPiece.GetValue(Tab[i])+TobEcr.GetValue(Tab[i]) ;
  TobPiece.PutValue(Tab[i],Tmp) ;
  end ;
TobPiece.PutValue('E_GENERAL' ,'Pièce ') ;
end ;

function VER_EURO.VerificationConvertJalPer(TobEcr: TOB): integer;
var i: integer ;
begin
Result:=0 ;
if TobEcr<>nil then begin
  for i:=1 to TobEcr.Detail.Count-1 do begin
    Result:=Result+VerificationConvertEcriture(TobEcr.Detail[i],False) ;
    Result:=Result+VerificationConvertPiece(TobEcr.Detail[i]);
    TobEccDeb:=nil ; TobEccCre:=nil ;
    end ;
  Result:=Result+VerificationConvertEcriture(TobPiece,True) ; //verifie la dernière piece !
  end ;
end ;

procedure VER_EURO.VerificationConvertCompta();
var Q: TQuery; TobCpta: TOB ;i: integer ;LeJal,LaPer: string ;
begin
//découpage de la table écriture en journal/periode
Q:=OpenSql('SELECT E_JOURNAL,E_PERIODE,J_MODESAISIE FROM ECRITURE'+
            'LEFT JOIN JOURNAL ON E_JOURNAL=J_JOURNAL GROUP BY E_JOURNAL,E_PERIODE,J_MODESAISIE',True) ;
try
TobJalPer.LoadDetailDB('$ECRITURE','','',Q,True) ;
finally
ferme(Q) ;
end ;
for i:=0 to TobJalPer.Detail.Count-1 do begin
  TobCpta:=nil ;
  LeJal:=TobJalPer.Detail[i].GetValue('E_JOURNAL') ;
  LaPer:=TobJalPer.Detail[i].GetValue('E_PERIODE') ;
  LaSai:=TobJalPer.Detail[i].GetValue('J_MODESAISIE') ;
  Q:=OpenSql('SELECT * FROM ECRITURE WHERE E_JOURNAL="'+LeJal+'" '+'AND E_PERIODE="'+LaPer+'" ORDER BY E_NUMEROPIECE',True) ;
  try
  TobCpta:=Tob.Create('$ECRITURE',nil,-1) ;
  TobCpta.LoadDetailDB('ECRITURE','','',Q,True) ;
  finally
  ferme(Q) ;
  end ;
  if VerificationConvertJalPer(TobCpta)>0 then begin
    TobCpta.InsertOrUpdateDB(true) ;
//    AfficheListe(TobEcart,'Liste des ecarts de conversions',[]);
    end ;
  if TobCpta<>nil then begin TobCpta.Free ; TobCpta:=nil ; end ;
  if TobEcart<>nil then begin TobEcart.Free ; TobEcart:=nil ; end ;
  end ;
end ;

procedure VER_EURO.InitValeur ;
begin
TobQuotite.LoadDetailDB('DEVISE','','',nil,False) ;
{$IFDEF SPEC302}
Q:=OpenSQL('SELECT SO_ECCEUROCREDIT,SO_ECCEURODEBIT,SO_OUVREPERTE, SO_OUVREBEN FROM SOCIETE',TRUE) ;
try
if not Q.Eof then begin
  EccEuroCredit:=Q.Fields[0].AsString ;
  EccEuroDebit :=Q.Fields[1].AsString ;
  Perte        :=Q.Fields[2].AsString ;
  Benef        :=Q.Fields[3].AsString ;
  end ;
finally
  ferme(Q) ;
end ;
{$ELSE}
EccEuroCredit:=GetParamSocSecur('SO_ECCEURODEBIT','') ;
EccEuroDebit :=GetParamSocSecur('SO_ECCEUROCREDIT','') ;
Perte        :=GetParamSocSecur('SO_OUVREPERTE','') ;
Benef        :=GetParamSocSecur('SO_OUVREBEN','') ;
{$ENDIF}
end ;

constructor VER_EURO.Create;
begin
  inherited Create;
  TobJalPer:=Tob.Create('$ECRITURE',nil,-1) ;
  TobErreur:=Tob.Create('$ERREUR',nil,-1) ;
  TobQuotite:=Tob.Create('DEVISE',nil,-1) ;
  TobEccDeb:=nil ; TobEccCre:=nil ;
  InitValeur ;
end ;

destructor VER_EURO.Destroy;
begin
  TobJalPer.Free ;
  TobErreur.Free ;
  TobQuotite.Free ;
  if TobPiece<>nil then TobPiece.Free ;
  inherited Destroy;
end ;

procedure VER_EURO.execute ;
begin
ParamOk:=CheckParam ;
if ParamOk then VerificationConvertCompta ;
end ;

procedure VER_EURO.affiche ;
begin
if (TobErreur<>nil) and ParamOk then
  if TobErreur.Detail.Count<=0 then PgiBox('Aucune erreur n''à été trouvée','Vérification Euro') ;
                               //else AfficheListe(TobErreur,'Liste des erreurs ',[]);
end ;

function VER_EURO.ChangePiece(TobEcr: TOB): boolean ;
begin
Result:=False ;
if TobPiece.GetValue('E_NUMEROPIECE')<>TobEcr.GetValue('E_NUMEROPIECE') then result:=True ;
// Saisie Bordereau
if LaSai='BOR' then begin
  if TobPiece.GetValue('E_DATECOMPTABLE')<>TobEcr.GetValue('E_DATECOMPTABLE') then Result:=True ;
  exit ;
  end ;
end ;

function VER_EURO.CheckCompte(Compte: string): integer ;
begin
Result:=0 ;
if Compte=EccEuroDebit  then Result:=-1 ;
if Compte=EccEuroCredit then Result:=1 ;
if Compte=Perte         then Result:=-1 ;
if Compte=Benef         then Result:=1 ;
end ;

function VER_EURO.CheckParam: boolean;
begin
Result:=False ;
if (Exo<>VH^.EnCours.Code) and (Exo<>VH^.Suivant.Code) then begin
  PgiBox('Exercice cloturé traitement impossible','Vérification Euro') ;
  Exit ;
  end ;
Result:=True ;
end ;


procedure VER_EURO.AjoutErreur(General: string;Err: integer ;F,E,D,Tmp,Taux,Quotite: double) ;
var LaTob: Tob ;  Tab: array[1..6] of string ;
begin
Tab[1]:='Convertion Francs => Euro   (Saisie Francs)' ; Tab[2]:='Convertion Francs => Devise (Saisie Francs)' ;
Tab[3]:='Convertion Euro   => Francs (Saisie Euro)' ;   Tab[4]:='Convertion Euro   => Devise (Saisie Euro)' ;
Tab[5]:='Convertion Euro   => Devise (Saisie Devise)' ; Tab[6]:='Convertion Euro   => Francs (Saisie Devise)' ;
if TobErreur=nil then Exit ;
if Err<0 then Err:=Err*-1 ;
LaTob:=Tob.Create('$ERREUR',TobErreur,-1) ;
LaTob.AddChampSup('General' ,True) ; LaTob.PutValue('General',General) ;
LaTob.AddChampSup('Erreur' ,True) ;  LaTob.PutValue('Erreur',Tab[Err]) ;
LaTob.AddChampSup('Francs' ,True) ;  LaTob.PutValue('Francs',F) ;
LaTob.AddChampSup('Euro'   ,True) ;  LaTob.PutValue('Euro',E) ;
LaTob.AddChampSup('Devise' ,True) ;  LaTob.PutValue('Devise',D) ;
LaTob.AddChampSup('Convert',True) ;  LaTob.PutValue('Convert',Tmp) ;
LaTob.AddChampSup('Taux',True)    ;  LaTob.PutValue('Taux',Taux) ;
LaTob.AddChampSup('Quotite',True) ;  LaTob.PutValue('Quotite',Quotite) ;
end ;
{*


   TOF VerEuro :


*}
procedure TOF_VerEuro.OnLoad ;
begin
Fexo:=THValComboBox(GetControl('CEXO')) ;
FEtab:=THValComboBox(GetControl('FETAB')) ;
FNatJnl1:=THValComboBox(GetControl('FNATJNL1')) ;
FNatJnl2:=THValComboBox(GetControl('FNATJNL2')) ;
SaufJnl:=THEdit(GetControl('SAUFJAL')) ;
FDateEcr1:=THEdit(GetControl('FDATEECR1')) ;
FDateEcr2:=THEdit(GetControl('FDATEECR2')) ;
SaufJnl.OnElipsisClick:=OnClickElipsis ;
FExo.ItemIndex:=0 ;
FEtab.ItemIndex:=0 ;
FNatJnl1.ItemIndex:=0 ;
FNatJnl2.ItemIndex:=FNatJnl2.Items.Count-1 ;
end ;

procedure TOF_VerEuro.OnUpdate ;
var Verif: VER_EURO ;
begin
Verif:=VER_EURO.Create ;
try
  Verif.Exo:=FExo.Value ;
  Verif.Etab:=FEtab.Value ;
  Verif.Jal1:=FNatJnl1.Value ;
  Verif.Jal2:=FNatJnl2.Value ;
  Verif.SaufJal:=SaufJnl.Text ;
  Verif.DateDeb:=StrToDate(FDateEcr1.Text) ;
  Verif.DateFin:=StrToDate(FDateEcr2.Text) ;
  Verif.execute ;
  Verif.affiche ;
finally
  Verif.Free ;
end ;
end ;

procedure TOF_VerEuro.OnClose ;
begin
//
end ;

procedure TOF_VerEuro.OnClickElipsis(Sender : TObject) ;
begin SaufJnl.Text:=AGLLanceFiche('CP','EXPS3JAL','','','') ; end ;

Initialization
registerclasses([TOF_VerEuro]);
end.
