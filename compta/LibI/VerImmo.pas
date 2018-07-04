unit VerImmo;

// CA - 07/07/1999 - Contrôle de l'état . On force à OUV si pas OK.

interface

uses
  (*Windows,Messages,Graphics, Dialogs,Menus ,Grids,StdCtrls,Buttons, ExtCtrls,*)
  SysUtils, Classes,  Controls, Forms, hmsgbox, HSysMenu, Hctrls, ImEnt, HEnt1, ImEdCalc,Outils
  {$IFDEF eAGLClient}
  ,UTob
  {$ELSE}
  {$IFDEF ODBCDAC}
  , odbcconnection, odbctable, odbcquery, odbcdac
  {$ELSE}
    {$IFNDEF DBXPRESS}
    ,dbtables
    {$ELSE}
    ,uDbxDataSet
    {$ENDIF}
    {$ENDIF}
   {$ENDIF}
  {$IFDEF SERIE1}
  ;
  {$ELSE}
  , RapSuppr ;
  {$ENDIF}

const tvImmos = 1 ;
      tvCptes = 2 ;
      tvAmort = 3 ;
      tvOp    = 4 ;
      tvEch   = 5 ;

type
  TFVerImmo = class(TForm)
    HMTrad: THSystemMenu;
    Msg: THMsgBox;
    MsgRien: THMsgBox;
    MsgLib: THMsgBox;
  private
    {$IFDEF SERIE1}
    LaListe : TStringList;
    {$ELSE}
    LaListe : TList ;
    {$ENDIF SERIE1}
    NbError : integer ;
    OkVerif : boolean ;
    Q : TQuery ;
    procedure SQLImmos ;
    procedure SQLCptes ;
    procedure SQLAmort ;
    procedure SQLOp ;
    procedure SQLEch ;
    {$IFDEF SERIE1}
    function FillListe(Code, Lib, Rem : string ; i : byte) : string ;
    {$ELSE}
    function FillListe(Code, Lib, Rem : string ; i : byte) : DelInfo ;
    {$ENDIF SERIE1}
    function VerifCode(Code, Lib : string) : boolean ;
    function VerifLibelle(Code, SQLLib : string) : Boolean ;
    function VerifCompteGene(Code, Lib, SQLCompte : string) : boolean ;
    function VerifASCII(Code, Lib, SQLCompte : string) : boolean ;
    procedure CorrigeASCII(var Code : string) ;
    function IsNature(Nature : string) : boolean ;
    function VerifNature(Code, Lib, SQLNature : string) : boolean ;
    //function VerifJournal(Code, Lib, SQLJournal : string) : boolean ;
    function VerifTiers(Code, Lib, SQLTiers : string) : boolean ;
    function IsMethode(Methode : string) : boolean ;
    function VerifMethode(Code, Lib, SQLMethode : string) : boolean ;
    //function VerifQualif(Code, Lib, SQLQualif : string) : boolean ;
    function VerifMontant(Code, Lib, SQLMontant : string) : boolean ;
    function VerifCodeImmo(Code, Lib : string) : boolean ;
    function VerifOperation(Code, Lib, SQLOperation : string) : boolean ;
    function VerifEtat(SQLCode, SQLLib: string): Boolean;
    function VerifLogAcquisition(Code,lib : string) : boolean ;
    procedure ModifListe(s: string) ;
  public
    bMaj : boolean ;
    Tv : integer ;
    procedure Execute ;
  end;

{$IFDEF SERIE1}
procedure ExecuteVerifImmos(tVerif : integer; bRepare : boolean;var Liste:TstringList) ;
{$ELSE}
procedure ExecuteVerifImmos(tVerif : integer; bRepare : boolean) ;
{$ENDIF SERIE1}

implementation

uses ImPlan, ImOuPlan
{$IFDEF SERIE1}
{$ELSE}
  , IMMO_TOM
{$ENDIF}
;

{$R *.DFM}

{$IFDEF SERIE1} //YCP 01/12/00
procedure ExecuteVerifImmos(tVerif : integer; bRepare : boolean;var Liste:TstringList) ;
var Verif : TFVerImmo ;
begin
  Verif:=TFVerImmo.Create(Application) ;
  Verif.bMaj:=bRepare ;
  Verif.Tv:=tVerif ;
  Verif.LaListe:=Liste ;
  Verif.Execute ;
  Verif.Free ;
end ;
{$ELSE}
procedure ExecuteVerifImmos(tVerif : integer; bRepare : boolean) ;
var Verif : TFVerImmo ;
begin
  Verif:=TFVerImmo.Create(Application) ;
  Verif.bMaj:=bRepare ;
  Verif.Tv:=tVerif ;
  Verif.Execute ;
  Verif.Free ;
end ;
{$ENDIF}

procedure TFVerImmo.Execute ;
begin
  OkVerif:=TRUE ; NbError:=0 ;
  {$IFDEF SERIE1}
  {$ELSE}
  LaListe:=TList.Create ;
  {$ENDIF}
  if Tv=tvImmos then SQLImmos ;
  if Tv=tvCptes then SQLCptes ;
  if Tv=tvAmort then SQLAmort ;
  if Tv=tvOp    then SQLOp ;
  if Tv=tvEch   then SQLEch ;
  {$IFDEF SERIE1}
  {$ELSE}
  if not OkVerif then
     begin
     RapportdErreurMvt(Laliste,77,bMaj,False);
//    if bMaj then ExecuteVerifImmos(Tv,True);
     end
  else MsgRien.Execute(Tv,'','') ; // sinon message tout est ok
  LaListe.Free ;
  {$ENDIF}
end ;

procedure TFVerImmo.SQLImmos ;
var Code, Lib, Nature,MethodeEco, Etat : string ;
begin
  Q:=OpenSQL('SELECT * FROM IMMO ORDER BY I_IMMO', not bMaj) ;
  while not Q.EOF do
  begin
    Code      :=Q.FindField('I_IMMO').AsString ;
    Lib       :=Q.FindField('I_LIBELLE').AsString ;
    Nature    :=Q.FindField('I_NATUREIMMO').AsString;
    MethodeEco:=Q.FindField('I_METHODEECO').AsString;
    Etat      :=Q.FindField('I_ETAT').AsString;
    if not VerifCode(Code, Lib) then OkVerif:=FALSE ;
    if not VerifLibelle(Code,'I_LIBELLE') then OkVerif:=FALSE ;
    if not VerifEtat('I_IMMO','I_ETAT') then OkVerif:=FALSE ;
    if not VerifNature(Code,MsgLib.Mess[0], 'I_NATUREIMMO') then OkVerif:=FALSE ;
    if Nature <> 'LOC' then
      if not  VerifCompteGene(Code,    MsgLib.Mess[1], 'I_COMPTEIMMO') then OkVerif:=FALSE ;
    if (Nature = 'LOC') or (Nature='CB') then
    begin
      if not  VerifCompteGene(Code,MsgLib.Mess[2], 'I_COMPTELIE') then OkVerif:=FALSE ;
      if not  VerifTiers(Code,MsgLib.Mess[20], 'I_ORGANISMECB')   then OkVerif:=FALSE ;
    end;
    if Nature='PRO' then
    begin
      if not VerifCompteGene(Code,MsgLib.Mess[3], 'I_COMPTEAMORT')    then OkVerif:=FALSE ;
      if not VerifCompteGene(Code,MsgLib.Mess[4], 'I_COMPTEDOTATION') then OkVerif:=FALSE ;
      if not VerifCompteGene(Code,MsgLib.Mess[5], 'I_COMPTEDEROG')    then OkVerif:=FALSE ;
      if not VerifCompteGene(Code,MsgLib.Mess[6], 'I_REPRISEDEROG')   then OkVerif:=FALSE ;
      if not VerifCompteGene(Code,MsgLib.Mess[7], 'I_PROVISDEROG')    then OkVerif:=FALSE ;
      if not VerifCompteGene(Code,MsgLib.Mess[8], 'I_DOTATIONEXC')    then OkVerif:=FALSE ;
      if not VerifCompteGene(Code,MsgLib.Mess[9], 'I_VAOACEDEE')      then OkVerif:=FALSE ;
      if not VerifCompteGene(Code,MsgLib.Mess[10],'I_VACEDEE')        then OkVerif:=FALSE ;
      if not VerifCompteGene(Code,MsgLib.Mess[11],'I_AMORTCEDE')      then OkVerif:=FALSE ;
      if not VerifCompteGene(Code,MsgLib.Mess[12],'I_REPEXPLOIT')     then OkVerif:=FALSE ;
      if not VerifCompteGene(Code,MsgLib.Mess[13],'I_REPEXCEP')       then OkVerif:=FALSE ;
    end;
    if ((Nature = 'PRO') or (Nature = 'CB')) and (Etat <> 'FER') then
    begin
      if not VerifMethode(Code,MsgLib.Mess[14],  'I_METHODEECO')       then OkVerif:=FALSE ;
    end;
    if not VerifMontant(Code,MsgLib.Mess[19],   'I_MONTANTHT') then OkVerif:=FALSE ;
    if not VerifOperation (Code,MsgLib.Mess[21],'I_OPERATION') then OkVerif:=FALSE ;
    if not VerifLogAcquisition(Code,lib)                       then OkVerif:=FALSE ;
    Q.Next ;
  end ;
  Ferme(Q) ;
end ;

procedure TFVerImmo.SQLCptes ;
var Code,prefixe : string ;
begin
  Q:=OpenSQL('SELECT * FROM IMMOCPTE', not bMaj) ;
  {$IFDEF SERIE1}
  prefixe:='IC' ;
  {$ELSE}
  prefixe:='PC' ;
  {$ENDIF}
  while not Q.EOF do
  begin
    Code:=Q.FindField(prefixe+'_COMPTEIMMO').AsString ;
    if not VerifCompteGene(Code, MsgLib.Mess[1],  prefixe+'_COMPTEIMMO')     then OkVerif:=FALSE ;
    if not VerifCompteGene(Code, MsgLib.Mess[3],  prefixe+'_COMPTEAMORT')    then OkVerif:=FALSE ;
    if not VerifCompteGene(Code, MsgLib.Mess[4],  prefixe+'_COMPTEDOTATION') then OkVerif:=FALSE ;
    if not VerifCompteGene(Code, MsgLib.Mess[5],  prefixe+'_COMPTEDEROG')    then OkVerif:=FALSE ;
    if not VerifCompteGene(Code, MsgLib.Mess[6],  prefixe+'_REPRISEDEROG')   then OkVerif:=FALSE ;
    if not VerifCompteGene(Code, MsgLib.Mess[7],  prefixe+'_PROVISDEROG')    then OkVerif:=FALSE ;
    if not VerifCompteGene(Code, MsgLib.Mess[8],  prefixe+'_DOTATIONEXC')    then OkVerif:=FALSE ;
    if not VerifCompteGene(Code, MsgLib.Mess[9],  prefixe+'_VOACEDE')        then OkVerif:=FALSE ;
    if not VerifCompteGene(Code, MsgLib.Mess[10 ],prefixe+'_VACEDEE')        then OkVerif:=FALSE ;
    if not VerifCompteGene(Code, MsgLib.Mess[11], prefixe+'_AMORTCEDE')      then OkVerif:=FALSE ;
    if not VerifCompteGene(Code, MsgLib.Mess[12], prefixe+'_REPEXPLOIT')     then OkVerif:=FALSE ;
    if not VerifCompteGene(Code, MsgLib.Mess[13], prefixe+'_REPEXCEP')       then OkVerif:=FALSE ;
    Q.Next ;
  end ;
  Ferme(Q) ;
end ;

procedure TFVerImmo.SQLAmort ;
var Code : string ;
begin
  Q:=OpenSQL('SELECT * FROM IMMOAMOR ORDER BY IA_DATE ASC', not bMaj) ;
  while not Q.EOF do
    begin
    Code:=Q.FindField('IA_IMMO').AsString ;
    if not VerifCode(Code, '') then OkVerif:=FALSE ;
    if not VerifCodeImmo(Code, '') then OkVerif:=FALSE ;
    if not VerifCompteGene(Code, MsgLib.Mess[0], 'IA_COMPTEIMMO') then OkVerif:=FALSE ;
    Q.Next ;
  end ;
  Ferme(Q) ;
end ;

procedure TFVerImmo.SQLOp ;
var Code : string ;
begin
  Q:=OpenSQL('SELECT * FROM IMMOLOG', not bMaj) ;
  while not Q.EOF do
  begin
    Code:=Q.FindField('IL_IMMO').AsString ;
    if not VerifCode(Code, '') then OkVerif:=FALSE ;
    if not VerifCodeImmo(Code, '') then OkVerif:=FALSE ;
    Q.Next ;
  end ;
  Ferme(Q) ;
end ;

procedure TFVerImmo.SQLEch ;
var Code : string ;
begin
  Q:=OpenSQL('SELECT * FROM IMMOECHE', not bMaj) ;
  while not Q.EOF do
  begin
    Code:=Q.FindField('IH_IMMO').AsString ;
    if not VerifCode(Code, '') then OkVerif:=FALSE ;
    if not VerifCodeImmo(Code, '') then OkVerif:=FALSE ;
    Q.Next ;
  end ;
  Ferme(Q) ;
end ;

{$IFDEF SERIE1}
function TFVerImmo.FillListe(Code, Lib, Rem : string ; i : byte) : string ;
begin
  Inc(NbError) ;
  result:=Msg.Mess[i]+' '+Rem+' '+Code+' '+Lib ;
end ;
{$ELSE}
function TFVerImmo.FillListe(Code, Lib, Rem : string ; i : byte) : DelInfo ;
var X : DelInfo ;
begin
  Inc(NbError) ;
  X:=DelInfo.Create ;
  X.LeLib:=Lib ;
  X.LeMess2:=Msg.Mess[i]+' '+Rem ;
  X.LeCod:=Code ;
  Result:=X ;
end ;
{$ENDIF}

procedure TFVerImmo.ModifListe(s: string) ;
{$IFDEF SERIE1}
begin
  LaListe.Strings[LaListe.count-1]:=LaListe.Strings[LaListe.count-1]+s ;
end ;
{$ELSE}
var X : DelInfo ;
begin
  X:=LaListe.Items[LaListe.count-1];
  X.LeMess2:=X.LeMess2+s
end ;
{$ENDIF}

function TFVerImmo.VerifCode(Code, Lib : string) : boolean ;
begin
Result:=(Code<>'')  ;
if not result then
   LaListe.Add(FillListe(Code, Lib, '', 0));
end ;

function TFVerImmo.VerifLibelle(Code, SQLLib : string) : Boolean ;
var Lib : string ;
begin
Lib:=Q.FindField(SQLLib).AsString ;
Result:=(Lib<>'') ;
if not result then
   begin
   LaListe.Add(FillListe(Code, Lib, '', 1));
   if bMaj then
      begin
      Q.Edit ;
      Q.FindField(SQLLib).AsString:=code ;
      ModifListe(' : Corrigé!') ;
      Q.Post ;
      end ;
   end ;
END ;

function TFVerImmo.VerifEtat(SQLCode, SQLLib : string) : Boolean ;
var Code, Lib : string ;
begin
Code:=Q.FindField(SQLCode).AsString ;
Lib:=Q.FindField(SQLLib).AsString ;
Result:=(Lib<>'') ;
if not result then
   begin
   LaListe.Add(FillListe(Code, Lib, '', 13));
   if bMaj then
      begin
      Q.Edit ;
      Q.FindField(SQLLib).AsString:='OUV';
      ModifListe(' : Corrigé!') ;
      Q.Post ;
      end ;
   end ;
END ;

function TFVerImmo.VerifCompteGene(Code, Lib, SQLCompte : string) : boolean ;
var Compte : string ;
begin
Compte:=Q.FindField(SQLCompte).AsString ;
if Compte='' then  // Compte à blanc
   begin
   LaListe.Add(FillListe(Code, Lib, '', 0));
   result:=false ;
   end else
   begin
   if VHImmo^.Cpta[ImGeneTofb].lg<>Length(Compte) then  // Compte non 'bourré'
      begin
      LaListe.Add(FillListe(Code, Lib, '('+IntToStr(Length(Compte))+' car.)', 2));
      if bMaj then
         begin
         Q.Edit;
         if Length(Compte) > VHImmo^.Cpta[ImGeneTofb].lg then
            begin
            Q.FindField(SQLCompte).AsString := Copy (Compte,1,VHImmo^.Cpta[ImGeneTofb].lg)
            end else
            begin
            while Length(Compte) < VHImmo^.Cpta[ImGeneTofb].lg do Compte := Compte+'0';
            Q.FindField(SQLCompte).AsString := Compte;
            end;
         Q.Post;
         ModifListe(' : Corrigé!') ;
         end;
      end ;
   if not Presence('GENERAUX','G_GENERAL',Compte) then  // Existence du compte
      begin
      LaListe.Add(FillListe(Code, Lib, ' : '+Compte, 11));
      result:=false;
      end else
      begin
      Result:=VerifASCII(Code, Lib, SQLCompte) ;
      end ;
   end;
end ;

function TFVerImmo.VerifASCII(Code, Lib, SQLCompte : string) : boolean ;
var i : byte ; bFind : boolean; Compte : string ;
begin
bFind:=FALSE ;
Compte:=Q.FindField(SQLCompte).AsString ;
for i:=1 to Length(Compte) do
    begin
    if not (Compte[i] in ['0'..'9','A'..'z',VHImmo^.Cpta[ImGeneTofb].Cb,' ','.','/',':','-']) then
       begin
       bFind:=TRUE ;
       break ;
       end ;
    end ;

if bFind then
   begin
   LaListe.Add(FillListe(Code, Lib, Compte, 3)) ;
   if bMaj then
      begin
      CorrigeASCII(Compte) ;
      Q.Edit ;
      Q.FindField(SQLCompte).AsString:=Compte ;
      Q.Post ;
      ModifListe(' : Corrigé!') ;
      bFind:=FALSE ;
      end ;
   end ;
Result:=not bFind ;

end ;

procedure TFVerImmo.CorrigeASCII(var Code : string) ;
var i : byte ;
begin
for i:=1 to Length(Code) do
   begin
   if not (Code[i] in ['0'..'9','A'..'z',VHImmo^.Cpta[ImGeneTofb].Cb,' ','.','/',':','-']) then Code[i]:=VHImmo^.Cpta[ImGeneTofb].Cb ;
   end ;
end ;

function TFVerImmo.IsNature(Nature : string) : boolean ;
var Q : TQuery ;
begin
  Result:=FALSE ;
  Q:=OpenSql('SELECT CO_CODE FROM COMMUN WHERE CO_TYPE="NIM"', TRUE) ;
  while not Q.Eof do
     begin
     if Nature=Q.Fields[0].AsString then
        begin
        Result:=TRUE;
        break ;
        end ;
     Q.Next ;
  end ;
  Ferme(Q) ;
end ;

function TFVerImmo.VerifNature(Code, Lib, SQLNature : string) : boolean ;
var Nature : string ;
begin
Nature:=Q.FindField(SQLNature).AsString ;
Result:=(Nature<>'') or  IsNature(Nature) ;
if not result then
   begin
   LaListe.Add(FillListe(Code, Lib, '', 0));

   if bMaj then
      begin
      Q.Edit ;
      Q.FindField(SQLNature).AsString:='PRO' ;
      Q.Post ;
      ModifListe(' : Corrigé!') ;
      end ;
   end ;
end ;

(*function TFVerImmo.VerifJournal(Code, Lib, SQLJournal : string) : boolean ;
var Journal : string ;
begin
  Result:=TRUE ;
  Journal:=Q.FindField(SQLJournal).AsString ;
  if (Journal='') or not presence ('JOURNAL', 'J_JOURNAL',Journal) then
  begin
    if not bMaj then
    begin
      LaListe.Add(FillListe(Code, Lib, '', 0));
      Result:=FALSE ;
    end;
  end ;
end ;*)

function TFVerImmo.VerifTiers(Code, Lib, SQLTiers : string) : boolean ;
begin
  Result:=presence ('TIERS', 'T_AUXILIAIRE',Q.FindField(SQLTiers).AsString) ;
  if not result  then LaListe.Add(FillListe(Code, Lib, '', 0));
end ;

function TFVerImmo.IsMethode(Methode : string) : boolean ;
var Q : TQuery ;
begin
Result:=FALSE ;
Q:=OpenSql('SELECT CO_CODE FROM COMMUN WHERE CO_TYPE="MIM"', TRUE) ;
while not Q.Eof do
   begin
   if Methode=Q.Fields[0].AsString then
      begin
      Result:=TRUE;
      break ;
      end ;
   Q.Next ;
   end ;
Ferme(Q) ;
end ;

function TFVerImmo.VerifMethode(Code, Lib, SQLMethode : string) : boolean ;
var Methode : string ;
begin
  Methode:=Q.FindField(SQLMethode).AsString ;
  Result:=(Methode<>'') or IsMethode(Methode) ;
  if not result then LaListe.Add(FillListe(Code, Lib, '', 4));
end ;


function TFVerImmo.VerifMontant(Code, Lib, SQLMontant : string) : boolean ;
begin
Result:=Q.FindField(SQLMontant).AsFloat>=0 ;
if not result then
   begin
   LaListe.Add(FillListe(Code, Lib, '', 8));
   if bMaj then
      begin
      Q.Edit ;
      Q.FindField(SQLMontant).AsFloat:=0;
      Q.Post ;
      ModifListe(' : Corrigé!') ;
      end;
   end ;
end ;


function TFVerImmo.VerifCodeImmo(Code, Lib : string) : boolean ;
begin
  Result:=presence ('IMMO', 'I_IMMO',Code) ;
  if not result then
     LaListe.Add(FillListe(Code, Lib, '', 10));
end ;

function TFVerImmo.VerifOperation(Code, Lib, SQLOperation : string) : boolean ;
var
   stOperation: string ;
begin
stOperation:=Q.FindField(SQLOperation).AsString ;
Result:=(stOperation='-') or (stOperation='X') ;
if not result then
   begin
   LaListe.Add(FillListe(Code, Lib, '', 12));

   if bMaj then
      begin
      Q.Edit ;
      Q.FindField(SQLOperation).AsString:='-' ;
      Q.Post ;
      ModifListe(' : Corrigé!') ;
      end ;
   end ;
if ((Q.FindField ('I_OPEMUTATION').AsString = 'X') or
   (Q.FindField ('I_OPEECLATEMENT').AsString = 'X') or
   (Q.FindField ('I_OPECESSION').AsString = 'X') or
   (Q.FindField ('I_OPECHANGEPLAN').AsString = 'X') or
   (Q.FindField ('I_OPELIEUGEO').AsString = 'X') or
   (Q.FindField ('I_OPEETABLISSEMENT').AsString = 'X') or
   (Q.FindField ('I_OPELEVEEOPTION').AsString = 'X') or
   (Q.FindField ('I_OPEMODIFBASES').AsString = 'X')) and (stOperation = '-') then
   begin
   LaListe.Add(FillListe(Code, Lib, '', 12));
   Result:=FALSE ;
   if bMaj then
      begin
      Q.Edit ;
      Q.FindField(SQLOperation).AsString:='X' ;
      Q.Post ;
      ModifListe(' : Corrigé!') ;
      end ;
   end ;
end ;


function TFVerImmo.VerifLogAcquisition(Code,Lib : string) : boolean ;
  //21/04/99 gestion enreg immolog type ope ACQquisition
  procedure EnregLogAcquisition(CodeImmo : string; DateOpe : TDateTime;PlanActif : integer ; infos : TInfoLog);
  var Query : TQuery ;
  begin
    // Création nouvel enregistrement ImmoLog
    Query:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+W_W+'"', FALSE) ;
    Query.Insert ;
    InitNew ( Query ); { FQ 12155 - CA - 03/06/2003 - Initialisation des zones de IMMOLOG par défaut }
    Query.FindField('IL_TYPEOP').AsString:='ACQ' ;
    Query.FindField('IL_IMMO').AsString:=CodeImmo;
    Query.FindField('IL_DATEOP').AsDateTime:=DateOpe;
    Query.FindField('IL_ORDRE').AsInteger:=1;
    Query.FindField('IL_ORDRESERIE').AsInteger:=1;
    Query.FindField('IL_PLANACTIFAV').AsInteger:=0;
    Query.FindField('IL_PLANACTIFAP').AsInteger:=PlanActif;
    Query.FindField('IL_LIBELLE').AsString:=RechDom('TIOPEAMOR', Query.FindField('IL_TYPEOP').AsString, FALSE)+' '+DateToStr(DateOpe) ;
    Query.FindField('IL_TYPEMODIF').AsString:=AffecteCommentaireOperation(Query.FindField('IL_TYPEOP').AsString);
    //EPZ 23/11/00
    Query.FindField('IL_TVARECUPERABLE').AsFloat:=Infos.TVARecuperable;
    Query.FindField('IL_TVARECUPEREE').AsFloat:=Infos.TVARecuperee;
    //EPZ 23/11/00
    Query.Post ;
    Ferme(Query) ;
  end;
  //21/04/99 gestion enreg immolog type ope ACQquisition
var
  infos: TInfoLog;
begin
Result:=existeSQL('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+Code+'" AND IL_TYPEOP in ("ACQ","MUT")') ;
if not result then
   begin
   LaListe.Add(FillListe(Code, Lib, '', 14));
   if bMaj then
      begin
      Infos.TVARecuperable:=Q.FindField('I_TVARECUPERABLE').asfloat ;
      Infos.TVARecuperee  :=Q.FindField('I_TVARECUPEREE').asfloat;
      {$IFDEF eAGLClient}
      {$ELSE}
      EnregLogAcquisition(Code,Q.FindField('I_DATEPIECEA').AsDateTime,Q.FindField('I_PLANACTIF').AsInteger,infos);
      {$ENDIF}
      Result:=TRUE ;
      ModifListe(' : Corrigé!') ;
      end ;
   end ;
end ;

end.
