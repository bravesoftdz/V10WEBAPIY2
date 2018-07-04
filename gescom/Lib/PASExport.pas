unit PASExport;

interface

uses HStatus,Hent1,HMsgBox,Controls,dbTables,NumConv,HCtrls,sysutils,M3FP,Utob,
     FactUtil,FactComm, EntGC, Ent1, FactTob ;

type TPASExport  = class
     FF : TextFile ;
     TobPiece, TobAdresseFac : TOB ;
     TobSociete : Tob ;
     TOBPiedBase : tob ;
     Okok,UneLigne : boolean ;
     Constructor Create ;
     Destructor Destroy ; override ;
     procedure LoadPiece (CleDoc : R_CleDoc; TobPiece : Tob) ;
     Procedure ChargeSociete ;
     Function ExportDoc ( LaPiece : string ; App : boolean ) : boolean ;
     Function ExportLigne (TobLigne : Tob ) : boolean ;
     Function ExportTaxes (TPB : Tob ) : boolean ;
     Function ExportEnreg(Enreg : string  ) : boolean ;
     Public
     FichierExport : String ;
     Patience : Boolean ;
     end ;

Function ZeroG ( St : String ; L : integer ) : String ;
Function PasStrFMontant(Mt : Double) : string ;
Function PasTauxTVA ( VenteAchat : String ; NumCat : integer ; TOBL : TOB ) : double;

Function PasRel_ExportDocument (LaPiece : string ) : boolean ;

implementation

Function PasRel_ExportDocument (LaPiece : string ) : boolean ;
var PAS :TPASExport ;
begin
PAS:=TPASExport.create ;
Pas.FichierExport:='PAS_Export.fac' ;
Pas.Patience:=true ;
PAS.ExportDoc (LaPiece,True);
PgiBox('Document exporté','Export Pas.Rel') ;
PAS.free;
Result:=True;
end ;

Function TPASExport.ExportDoc ( LaPiece : string ; App : boolean ) : boolean ;
Var io : integer ;
    Cledoc: R_CleDoc ;
BEGIN
Result:=False ; Okok:=True ; UneLigne:=False ;
StringToCleDoc(LaPiece,Cledoc) ;
{Gestion fichier}
if FichierExport='' then Exit ;
AssignFile(FF,FichierExport) ;
{$I-}
if FileExists(FichierExport)and App
   then Append(FF)
   else ReWrite(FF) ;
{$I+}
if IoResult<>0 then
   BEGIN {$I-} CloseFile(FF) ; {$I+} io:=ioResult ;PgiBox('Impossible d''écrire dans le fichier '+FichierExport,'Export Pas.Rel') ; Exit ; END ;
{Traitement export}
TOBPiece:=TOB.Create('PIECE',Nil,-1) ;
TOBAdresseFac:=TOB.Create('ADRESSES',Nil,-1) ;
TOBPiedbase:=TOB.Create('_PiedBase',Nil,-1) ;
try
   LoadPiece (CleDoc,TobPiece) ;
   ExportEnreg('INFOGENE') ;                                            
   ExportEnreg('MONDEVIS') ;
   ExportEnreg('INFOREGL') ;
   ExportEnreg('ADRSGFOU') ;
   ExportEnreg('ADRBEPAI') ;
   ExportEnreg('ADRACCLI') ;
   ExportEnreg('CONDPAIE') ;
   ExportEnreg('LIGFACAV') ;
   ExportEnreg('RECREDOC') ;
   ExportEnreg('PRECATVA') ;
   ExportEnreg('MNTTOTDO') ;

finally
   TobPiece.free ; TobPiece:=nil;
   TobAdresseFac.free ; TobAdresseFac:=nil;
   TOBPiedBase.free ; TOBPiedBase:=nil;
   CloseFile(FF) ;
end;
if Not UneLigne then
   BEGIN
   if not App then
      begin
      AssignFile(FF,FichierExport) ;
      {$i-} Erase(FF) ; {$i+} io:=ioResult ;
      end;
   Okok:=False ; {PgiBox('Aucun élément à exporter','Export Pas.Rel') ; }
   END else
   BEGIN
   if (Not Okok)then BEGIN if PgiAsk('Il y a eu des anomalies, voulez vous continuer ?','Export Pas.Rel')=mrYes then Okok:=True ; END ;
        {else  PgiBox('Document exporté','Export Pas.Rel') ; }
   END ;
Result:=Okok ;
END ;

Destructor TPASExport.Destroy ;
begin
TobSociete.free ;
inherited Destroy ;
end;

Constructor   TPASExport.create ;
begin
inherited create ;
Patience:=False ;
FichierExport:='PasRel_Export.fac' ;
ChargeSociete ;
end;


procedure TPASExport.LoadPiece (CleDoc : R_CleDoc; TobPiece : Tob) ;
Var Q : TQuery ;
    RefAdresse : string ;
BEGIN
// Lecture Pied
Q:=OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False),True) ;
TOBPiece.SelectDB('',Q) ;
Ferme(Q) ;
// Lecture Lignes
Q:=OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',True) ;
TOBPiece.LoadDetailDB('LIGNE','','',Q,False,False) ;
Ferme(Q) ;
// Bases Taxes
Q:=OpenSQL('SELECT * FROM PIEDBASE WHERE '+WherePiece(CleDoc,ttdPiedBase,False),True) ;
TOBPiedBase.loadDetailDB('PIEDBASE','','',Q,False,False) ;
Ferme(Q) ;
// adresse de facturation
RefAdresse:= TobPiece.GetValue('GP_NATUREPIECEG')+';'+TobPiece.GetValue('GP_SOUCHE')+';'+intToStr(TobPiece.GetValue('GP_NUMERO'))+';' ;
Q:=OpenSQL('SELECT * FROM ADRESSES WHERE ADR_TYPEADRESSE="PIE" and ADR_REFCODE="'+RefAdresse+'"',True) ;
TOBAdresseFac.SelectDB('',Q) ;
Ferme(Q) ;
TobAdresseFac.addChampSup('T_NIF',true) ;
TobAdresseFac.addChampSup('T_SIRET',true) ;
TobAdresseFac.addChampSup('T_LIBELLE',true) ;
Q:=OpenSql('SELECT T_SIRET,T_NIF,T_LIBELLE from TIERS where T_TIERS="'+TobPiece.Getvalue('GP_TIERS')+'"',true) ;
if not Q.eof then
   begin
   TobAdresseFac.putValue('T_NIF', Q.FindField('T_NIF').asString) ;
   TobAdresseFac.putValue('T_SIRET', Q.FindField('T_SIRET').asString) ;
   TobAdresseFac.putValue('T_LIBELLE', Q.FindField('T_LIBELLE').asString) ;
   end else
   begin
   TobAdresseFac.putValue('T_NIF', '') ;
   TobAdresseFac.putValue('T_SIRET', '') ;
   TobAdresseFac.putValue('T_LIBELLE', '') ;
   end;
Ferme(Q) ;
end ;

Function TPASExport.ExportEnreg(Enreg : string  ) : boolean ;
var St,StW : string ;
    i : integer ;
begin
if Enreg='INFOGENE' then
   begin
   St:=Enreg ;
   St:=St+'380'+Format_String(TOBPiece.Getvalue('GP_SOUCHE'),3)+ZeroG(Inttostr(TOBPiece.Getvalue('GP_NUMERO')),9)+FormateDate('yyyymmdd',TOBPiece.Getvalue('GP_DATEPIECE'))+Format_String('',3) ;
   St:=St+Format_String(TOBPiece.Getvalue('GP_TIERS'),17)+Format_String('',146) ;
   St:=St+Format_String(' ',350) ;
   WriteLn(FF,St) ;
   end else
if Enreg='INFOREGL' then
   begin
   stW:= 'Au Capital de '+ FloatToStrF (TOBSociete.Getvalue('SO_CAPITAL'),ffNumber,10,0)+' F(rancs)';
   St:=Enreg ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_NATUREJURIDIQUE'),70) ;
   St:=St+Format_String(stW,70) ;
   St:=St+Format_String(' ',350) ;
   WriteLn(FF,St) ;
   end else
if Enreg='ADRSGFOU' then
   begin
   St:=Enreg ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_SIRET'),17)+Format_String('100',3) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_LIBELLE'),175) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_ADRESSE1'),35) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_ADRESSE2'),35) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_ADRESSE3'),35) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_VILLE'),35) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_CODEPOSTAL'),9) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_PAYS'),3) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_RC'),35) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_NIF'),35) ;
   St:=St+Format_String('',35) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_APE'),35) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_CONTACT'),35) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_TELEPHONE'),25) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_FAX'),25) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_TELEX'),25) ;
   St:=St+Format_String(' ',350) ;
   WriteLn(FF,St) ;
   end else
if Enreg='ADRBEPAI' then
   begin
   St:=Enreg ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_SIRET'),17)+Format_String('100',3) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_LIBELLE'),175) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_ADRESSE1'),35) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_ADRESSE2'),35) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_ADRESSE3'),35) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_VILLE'),35) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_CODEPOSTAL'),9) ;
   St:=St+Format_String(TOBSociete.Getvalue('SO_PAYS'),3) ;
   St:=St+Format_String('',245) ;
   St:=St+Format_String(' ',350) ;
   WriteLn(FF,St) ;
   end else
if Enreg='ADRACCLI' then
   begin
   St:=Enreg ;
   St:=St+Format_String(TOBAdresseFac.Getvalue('T_SIRET'),17)+Format_String('100',3) ;
   // en attendant correction bug sur adresse
   // St:=St+Format_String(TOBAdresseFac.Getvalue('ADR_LIBELLE'),35) ;
   St:=St+Format_String(TOBAdresseFac.Getvalue('T_LIBELLE'),175) ;
   St:=St+Format_String(TOBAdresseFac.Getvalue('ADR_ADRESSE1'),35) ;
   St:=St+Format_String(TOBAdresseFac.Getvalue('ADR_ADRESSE2'),35) ;
   St:=St+Format_String(TOBAdresseFac.Getvalue('ADR_ADRESSE3'),35) ;
   St:=St+Format_String(TOBAdresseFac.Getvalue('ADR_VILLE'),35) ;
   St:=St+Format_String(TOBAdresseFac.Getvalue('ADR_CODEPOSTAL'),9) ;
   St:=St+Format_String(TOBAdresseFac.Getvalue('ADR_PAYS'),3) ;
   St:=St+Format_String(TOBAdresseFac.Getvalue('T_NIF'),35) ; // code NIF
   St:=St+Format_String(TOBAdresseFac.Getvalue('ADR_CONTACT'),35) ;
   St:=St+Format_String(TOBAdresseFac.Getvalue('ADR_TELEPHONE'),25) ;
   St:=St+Format_String(' ',25) ; // fax contact
   St:=St+Format_String(' ',350) ;
   WriteLn(FF,St) ;
   end else
if Enreg='MONDEVIS' then
   begin
   St:=Enreg ;
   St:=St+Format_String(TOBPiece.Getvalue('GP_DEVISE'),3) ;
   St:=St+Format_String(' ',350) ;
   WriteLn(FF,St) ;
   end else
if Enreg='CONDPAIE' then
   begin
   St:=Enreg ;
   St:=St+Format_String('20',3) ;
   St:=St+Format_String(' ',234) ;
   St:=St+Format_String(' ',350) ;
   WriteLn(FF,St) ;
   end else
if Enreg='LIGFACAV' then
   begin
   if TobPiece.detail.count>0 then
      BEGIN
      If Patience then InitMove(TobPiece.detail.Count-1,'') ;
      for i:=0 to TobPiece.detail.Count-1 do
          BEGIN
          If Patience then MoveCur(False) ;
          if Not ExportLigne(TobPiece.detail[i]) then Okok:=False ;
          END ;
      If Patience then FiniMove ;
      END ;
   end else
if Enreg='PRECATVA' then
   begin
   if TobPiece.detail.count>0 then
      BEGIN
      for i:=0 to TobPiedBase.detail.Count-1 do
          BEGIN
          ExportTaxes(TobPiedBase.detail[i]) ;
          END ;
      END ;
   end else
if (Enreg='RECREDOC') and TOBPiece.Getvalue('GP_TOTALREMISEDEV') <> 0.0 then
   begin
   St:=Enreg ;
   St:=St+Format_String('Remise',35) ;
   St:=St+ZeroG(PasStrFMontant (TOBPiece.Getvalue('GP_TOTALBASEREMDEV')),35) ;
   St:=St+ZeroG(PasStrFMontant (TOBPiece.Getvalue('GP_REMISEPIED')),8) ;
   St:=St+ZeroG(PasStrFMontant (TOBPiece.Getvalue('GP_TOTALREMISEDEV')),35) ;
   St:=St+Format_String(' ',350) ;
   WriteLn(FF,St) ;
   end else
if Enreg='MNTTOTDO' then
   begin
   St:=Enreg ;
   St:=St+ZeroG(PasStrFMontant (TOBPiece.Getvalue('GP_TOTALHTDEV')),35) ;
   St:=St+ZeroG(PasStrFMontant (TOBPiece.Getvalue('GP_TOTALTTCDEV')-TOBPiece.Getvalue('GP_TOTALHTDEV')),35) ;
   St:=St+ZeroG(PasStrFMontant (TOBPiece.Getvalue('GP_TOTALTTCDEV')),35) ;
   St:=St+ZeroG(PasStrFMontant (TOBPiece.Getvalue('GP_ACOMPTEDEV')),35) ;
   St:=St+ZeroG(PasStrFMontant (TOBPiece.Getvalue('GP_TOTALTTCDEV')-TOBPiece.Getvalue('GP_ACOMPTEDEV')),35) ;
   St:=St+Format_String('',161) ;
   St:=St+Format_String(' ',350) ;
   WriteLn(FF,St) ;
   end ; {else}
Result:=True ;
end;

Function TPASExport.ExportLigne (TobLigne : Tob ) : boolean ;
var St : string ;
begin
if TobLigne.Getvalue('GL_TYPELIGNE')='ART' then
   begin
   St:='LIGFACAV' ;
   St:=St+ZeroG(Inttostr(TobLigne.Getvalue('GL_NUMLIGNE')),6) ;
   St:=St+Format_String('',35)+Format_String(TobLigne.Getvalue('GL_ARTICLE'),35) ;
   St:=St+Format_String('',35)+'SA ' ;
   St:=St+Format_String('',35)+Format_String(TobLigne.Getvalue('GL_LIBELLE'),70) ;
   St:=St+Format_String('',8)+FormateDate('yyyymmdd',TobLigne.Getvalue('GL_DATEPIECE')) ;
   St:=St+Format_String('',186)+ZeroG(PasStrFMontant (TobLigne.Getvalue('GL_QTEFACT')),15) ;
   St:=St+Format_String('PCE',3)+ZeroG(PasStrFMontant (TobLigne.Getvalue('GL_PUHTDEV')),15) ;
   St:=St+ZeroG(PasStrFMontant (TobLigne.Getvalue('GL_PUHTNETDEV')),15) ;
   St:=St+Format_String('',12+35)+ZeroG(PasStrFMontant (TobLigne.Getvalue('GL_MONTANTHTDEV')+TobLigne.Getvalue('GL_TOTREMLIGNEDEV')),35) ;
   St:=St+Format_String('',78) ;
   St:=ST+Format_String(PasStrFMontant(PasTauxTVA('VEN',1,TobLigne)),17) ;
   St:=St+ZeroG(PasStrFMontant (TobLigne.Getvalue('GL_TOTALTAXEDEV1')),35) ;
   St:=St+Format_String('',175) ; {en attente}
   UneLigne:=True ;
   end else
if TobLigne.Getvalue('GL_TYPELIGNE')='COM' then
   begin
   if Uneligne then St:='TXTLIBR2' else St:='TXTLIBR1' ;
   St:=St+Format_String(TobLigne.Getvalue('GL_LIBELLE'),350) ;
   end else
if TobLigne.Getvalue('GL_TYPELIGNE')='TOT' then
   begin
   Result:=True ;
   exit;
   end;
St:=St+Format_String(' ',350) ;
WriteLn(FF,St) ;
Result:=True ;
end ;

Function PasTauxTVA ( VenteAchat : String ; NumCat : integer ; TOBL : TOB ) : double;
Var  TOBT : TOB ;
     RegimeTaxe,FamilleTaxe : String ;
BEGIN
Result:=0.0;

FamilleTaxe:=TOBL.GetValue('GL_FAMILLETAXE'+IntToStr(NumCat)) ;
RegimeTaxe:=TOBL.GetValue('GL_REGIMETAXE') ;
TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],['TX'+IntToStr(NumCat),RegimeTaxe,FamilleTaxe],False) ;
if TOBT<>Nil then
   BEGIN
   if VenteAchat='VEN' then Result:=TOBT.GetValue('TV_TAUXVTE') else Result:=TOBT.GetValue('TV_TAUXACH') ;
   END ;
end ;

Function TPASExport.ExportTaxes (TPB : Tob ) : boolean ;
var St : string ;
begin
St:='PRECATVA' ;
ST:=ST+ZeroG(PasStrFMontant (TPB.Getvalue('GPB_TAUXTAXE')),17) ;
St:=St+Format_String('',3) ;
ST:=ST+ZeroG(PasStrFMontant (TPB.Getvalue('GPB_BASEDEV')),15) ;
ST:=ST+ZeroG(PasStrFMontant (TPB.Getvalue('GPB_VALEURDEV')),35) ;
St:=St+Format_String(' ',333) ;
WriteLn(FF,St) ;
Result:=True ;
end;

Procedure TPASExport.ChargeSociete ;
var QQ:TQuery ;
begin
QQ:=OpenSql ('SELECT SO_SOCIETE,SO_LIBELLE,SO_ADRESSE1,SO_ADRESSE2,SO_ADRESSE3,SO_CODEPOSTAL'
            +',SO_VILLE,SO_DIVTERRIT,SO_PAYS,SO_TELEPHONE,SO_FAX,SO_TELEX,SO_NIF,SO_APE,SO_SIRET,SO_RC  '
            +',SO_NATUREJURIDIQUE,SO_CAPITAL'
            +' FROM SOCIETE where SO_SOCIETE="'+V_PGI.CodeSociete+'"', TRUE);
TobSociete:=Tob.create('SOCIETE',nil,-1) ;
TobSociete.selectDB ( '',QQ) ;
Ferme(QQ) ;
end ;


Function ZeroG ( St : String ; L : integer ) : String ;
BEGIN
St:=Copy(St,1,L) ;
While Length(St)<L do St:='0'+St ;
Result:=St ;
END ;

FUNCTION PasStrFMontant(Mt : Double) : string ;
Var i : integer ;
    st : string ;
BEGIN
st:=STRFMONTANT (Mt,15,2,'',False)  ;
if DecimalSeparator=',' then
   begin
   i:=Pos(',',St) ;
   if(i>0) then St[i]:='.';
   end;
Result:=St ;
END ;

procedure AGLPasRelExportDocument (parms:array of variant; nb: integer ) ;
Begin
PasRel_ExportDocument(string(Parms[0]));
End;


Initialization
RegisterAglProc('PasRelExportDocument',False,1,AGLPasRelExportDocument);


end.
