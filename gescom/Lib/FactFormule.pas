unit FactFormule;

interface
Uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
{$ENDIF}
     SysUtils, Dialogs, Utiltarif,TarifUtil, UtilPGI, UtilGC, AGLInit, EntGC, SaisUtil, Windows,
     Forms, Classes, AGLInitGC, UtilArticle , UtilDimArticle, HDimension, HMsgBox,
     ParamSoc, HPanel,UTilFonctionCalcul,uEntCommun;

function ChargerTobArt(TOBArt,TobFormule : Tob; VenteAchat,LaCle1 : String ; QQ : TQuery ; SansMemo : Boolean = FALSE) :boolean;
Procedure LoadLesFormules ( TOBPiece,TOBLFormule : TOB ; CleDoc : R_CleDoc ) ;
Procedure InsertFormuleToTOB (TOBLigne,TOBLFormule : TOB; RFonct : R_FonctCal) ;
//Procedure DeleteFormuleToTOB (TOBLigne,TOBLFormule : TOB) ;
Procedure UpdateFormuleToTOB (TOBLigne,TOBLFormule : TOB; RFonct : R_FonctCal) ;
//Procedure UpdateLesFormules (TOBL,TOBLF : TOB) ;
procedure ValideLesFormules (TOBPiece,TOBLFormule : TOB);
function DetruitLesFormules (TOBPiece,TOBLFormule : TOB) : Boolean;
procedure ReliquatFormule (TOBLigne,TOBLFormule : TOB) ;


implementation

function ChargerTobArt(TOBArt,TobFormule : Tob; VenteAchat,LaCle1 : String ; QQ : TQuery ; SansMemo : Boolean = FALSE) :boolean;
Var TOBADet : TOB;
    QQSuite : TQuery;
    VteAch : string;

begin
if (LaCle1<>'') or (QQ<>Nil) then
begin
	if LaCle1 <> '' then
  begin
    QQSuite := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
      						 '(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=GA_FOURNPRINC) AS LIBELLEFOU FROM ARTICLE A '+
                   'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                   'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+LaCle1+'"',true,-1, '', True);

  	Result:=TOBArt.SelectDB('',QQSuite,SansMemo);
    ferme (QQSuite);
  end else
  begin
  	Result:=TOBArt.SelectDB('',QQ,SansMemo)
  end;
end
else Result:=(TOBArt<>Nil);

if TOBArt <> nil then InitChampsSupArticle (TOBART);

TOBADet:=Nil;

  if      VenteAchat='VEN' then VteAch:='VTE'
  else if VenteAchat='ACH' then VteAch:='ACH'
  else    exit;

  if TobFormule= nil then exit;

  if Result then
  begin
    if TobFormule.FindFirst(['GAF_ARTICLE'],[TOBArt.GetValue('GA_ARTICLE')],False)<>Nil then Exit;
      QQSuite:=OpenSql('Select * from ARTICLEQTE where GAF_ARTICLE="' + TOBArt.GetValue('GA_ARTICLE')+'"',True,-1, '', True);
      if Not QQSuite.Eof then TOBADet:=TOB.CreateDB('ARTICLEQTE',TobFormule,-1,QQSuite) ;
      Ferme(QQSuite);
      if TOBADet<> Nil then
      begin
        QQSuite:=OpenSql('Select * from ARTFORMULEVAR where GAV_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" and GAV_VENTEACHAT="'+VteAch+'"',True,-1, '', True);
        if Not QQSuite.Eof then TOBADet.LoadDetailDB('ARTICLEQTE','','',QQSuite,False) ;
        Ferme(QQSuite);
      end;
  end;
end;

Procedure LoadLesFormules ( TOBPiece,TOBLFormule : TOB ; CleDoc : R_CleDoc ) ;
Var QQ : TQuery ;
    i_ind,NumL,NumLPrec,IndiceFormule : integer ;
    TOBVrac,TOBLigne,TOBLF,TobLigF : TOB ;
    St : string ;
BEGIN
TOBVrac:=TOB.Create('',Nil,-1) ;
St:=' GLF_NATUREPIECEG="'+CleDoc.NaturePiece+'" AND GLF_SOUCHE="'+CleDoc.Souche+'" '
   +' AND GLF_NUMERO='+IntToStr(CleDoc.NumeroPiece)+' AND GLF_INDICEG='+IntToStr(CleDoc.Indice)
   +' AND GLF_TYPEFORMULE="QTE" ' ;
QQ:=OpenSQL('SELECT * FROM LIGNEFORMULE WHERE '+St+
            ' ORDER BY GLF_NUMLIGNE DESC,GLF_TYPEFORMULE,GLF_RANG DESC',True,-1, '', True) ;
if Not QQ.EOF then TOBVrac.LoadDetailDB('LIGNEFORMULE','','',QQ,False) ;
Ferme(QQ) ; NumLPrec:=0 ;  TobLigF:=Nil;
for i_ind:=TOBVrac.Detail.Count-1 Downto 0 do
  begin
  TOBLF:=TOBVrac.Detail[i_ind] ;
  NumL:=TOBLF.GetValue('GLF_NUMLIGNE') ;
  if Numl<> NumLPrec then
    begin
    NumLPrec:=NumL;
    {recherche de la ligne de pièce concernée}
    TOBLigne:=TOBPiece.FindFirst(['GL_NUMLIGNE'],[NumL],False) ;
    if TOBLigne<>Nil then
      begin
      {Recherche du noeud des lots de la ligne, si pas ok le créer}
      TOBLF.ChangeParent(TOBLFormule,-1) ;
      TobLigF:=TOBLF;
      IndiceFormule:=TOBLFormule.Detail.Count ;
      TOBLigne.AddChampSupValeur('GL_INDICEFORMULE',IndiceFormule) ;
      end;
    end else
    begin
    if TobLigF<>Nil then TOBLF.ChangeParent(TobLigF,-1) ;
    end;
  end ;
TOBVrac.Free ;
//GCVoirTob(TOBLFormule);
end ;

Procedure InsertFormuleToTOB (TOBLigne,TOBLFormule : TOB; RFonct : R_FonctCal) ;
Var ind,IndiceFormule : integer;
    TOBLF,TOBL : TOB;
begin
TOBLF:=Tob.Create('LIGNEFORMULE',TOBLFormule,-1);
TOBLF.PutValue('GLF_NATUREPIECEG',TOBLigne.GetValue('GL_NATUREPIECEG'));
TOBLF.PutValue('GLF_SOUCHE',TOBLigne.GetValue('GL_SOUCHE'));
TOBLF.PutValue('GLF_NUMERO',TOBLigne.GetValue('GL_NUMERO'));
TOBLF.PutValue('GLF_INDICEG',TOBLigne.GetValue('GL_INDICEG'));
TOBLF.PutValue('GLF_NUMLIGNE',TOBLigne.GetValue('GL_NUMLIGNE'));
TOBLF.PutValue('GLF_TYPEFORMULE','QTE');
TOBLF.PutValue('GLF_RANG',0);
TOBLF.PutValue('GLF_ARTICLE',TOBLigne.GetValue('GL_ARTICLE'));
TOBLF.PutValue('GLF_FORMULE',RFonct.Formule);
TOBLF.PutValue('GLF_EXPRESSION',RFonct.Expression);
IndiceFormule:=TOBLFormule.Detail.Count ;
if TOBLigne.FieldExists('GL_INDICEFORMULE') then TOBLigne.DelChampSup('GL_INDICEFORMULE',False);
TOBLigne.AddChampSupValeur('GL_INDICEFORMULE',IndiceFormule) ;
for ind:=0 to High(RFonct.VarLibelle) do
  begin
  if RFonct.VarLibelle[ind]<>'' then
    begin
    TOBL:=Tob.Create('LIGNEFORMULE',TOBLF,-1);
    TOBL.PutValue('GLF_NATUREPIECEG',TOBLigne.GetValue('GL_NATUREPIECEG'));
    TOBL.PutValue('GLF_SOUCHE',TOBLigne.GetValue('GL_SOUCHE'));
    TOBL.PutValue('GLF_NUMERO',TOBLigne.GetValue('GL_NUMERO'));
    TOBL.PutValue('GLF_INDICEG',TOBLigne.GetValue('GL_INDICEG'));
    TOBL.PutValue('GLF_NUMLIGNE',TOBLigne.GetValue('GL_NUMLIGNE'));
    TOBL.PutValue('GLF_TYPEFORMULE','QTE');
    TOBL.PutValue('GLF_RANG',ind+1);
    TOBL.PutValue('GLF_ARTICLE',TOBLigne.GetValue('GL_ARTICLE'));
    TOBL.PutValue('GLF_FORMULE','');
    TOBL.PutValue('GLF_EXPRESSION','');
    TOBL.PutValue('GLF_LIBELLEVAR',RFonct.VarLibelle[ind]);
    TOBL.PutValue('GLF_VALEURVAR',RFonct.VarValeur[ind]);
    if RFonct.Affichable[ind] then TOBL.PutValue('GLF_VISIBLEVAR','X')
                              else TOBL.PutValue('GLF_VISIBLEVAR','-');
    if RFonct.Modifiable[ind] then TOBL.PutValue('GLF_ENABLEVAR','X')
                              else TOBL.PutValue('GLF_ENABLEVAR','-');
    end;
  end;
end;

Procedure DeleteFormuleToTOB (TOBLigne,TOBLFormule : TOB) ;
Var IndiceFormule : integer;
begin
if TOBLigne.FieldExists('GL_INDICEFORMULE') then
  IndiceFormule:=TOBLigne.GetValue('GL_INDICEFORMULE')
  else IndiceFormule:=0;
if IndiceFormule=0 then exit;
TOBLFormule.Detail[IndiceFormule-1].Free;
if TOBLigne.FieldExists('GL_INDICEFORMULE') then TOBLigne.DelChampSup('GL_INDICEFORMULE',False);
end;

Procedure UpdateFormuleToTOB (TOBLigne,TOBLFormule : TOB; RFonct : R_FonctCal) ;
Var ind,IndiceFormule : integer;
    TOBLF,TOBL : TOB;
begin
if TOBLigne.FieldExists('GL_INDICEFORMULE') then
  IndiceFormule:=TOBLigne.GetValue('GL_INDICEFORMULE')
  else IndiceFormule:=0;
if IndiceFormule=0 then exit;
TOBLF:=TOBLFormule.Detail[IndiceFormule-1];
TOBLF.PutValue('GLF_NATUREPIECEG',TOBLigne.GetValue('GL_NATUREPIECEG'));
TOBLF.PutValue('GLF_SOUCHE',TOBLigne.GetValue('GL_SOUCHE'));
TOBLF.PutValue('GLF_NUMERO',TOBLigne.GetValue('GL_NUMERO'));
TOBLF.PutValue('GLF_INDICEG',TOBLigne.GetValue('GL_INDICEG'));
TOBLF.PutValue('GLF_NUMLIGNE',TOBLigne.GetValue('GL_NUMLIGNE'));
TOBLF.PutValue('GLF_TYPEFORMULE','QTE');
TOBLF.PutValue('GLF_RANG',0);
TOBLF.PutValue('GLF_ARTICLE',TOBLigne.GetValue('GL_ARTICLE'));
TOBLF.PutValue('GLF_FORMULE',RFonct.Formule);
TOBLF.PutValue('GLF_EXPRESSION',RFonct.Expression);
for ind:=0 to High(RFonct.VarLibelle) do
  begin
  if RFonct.VarLibelle[ind]<>'' then
    begin
    if ind>=TOBLF.Detail.Count then TOBL:=Tob.Create('LIGNEFORMULE',TOBLF,-1)
                               else TOBL:=TOBLF.Detail[ind];
    TOBL.PutValue('GLF_NATUREPIECEG',TOBLigne.GetValue('GL_NATUREPIECEG'));
    TOBL.PutValue('GLF_SOUCHE',TOBLigne.GetValue('GL_SOUCHE'));
    TOBL.PutValue('GLF_NUMERO',TOBLigne.GetValue('GL_NUMERO'));
    TOBL.PutValue('GLF_INDICEG',TOBLigne.GetValue('GL_INDICEG'));
    TOBL.PutValue('GLF_NUMLIGNE',TOBLigne.GetValue('GL_NUMLIGNE'));
    TOBL.PutValue('GLF_TYPEFORMULE','QTE');
    TOBL.PutValue('GLF_RANG',ind+1);
    TOBL.PutValue('GLF_ARTICLE',TOBLigne.GetValue('GL_ARTICLE'));
    TOBL.PutValue('GLF_FORMULE','');
    TOBL.PutValue('GLF_EXPRESSION','');
    TOBL.PutValue('GLF_LIBELLEVAR',RFonct.VarLibelle[ind]);
    TOBL.PutValue('GLF_VALEURVAR',RFonct.VarValeur[ind]);
    if RFonct.Affichable[ind] then TOBL.PutValue('GLF_VISIBLEVAR','X')
                              else TOBL.PutValue('GLF_VISIBLEVAR','-');
    if RFonct.Modifiable[ind] then TOBL.PutValue('GLF_ENABLEVAR','X')
                              else TOBL.PutValue('GLF_ENABLEVAR','-');
    end;
  end;
end;

Procedure UpdateLesFormules (TOBL, TOBLF : TOB) ;
Var NatureG,Souche : String ;
    Numero,IndiceG,ind1,NumL : integer ;
    TOBDet   : TOB ;
begin
if ((TOBL=Nil) or (TOBLF=Nil)) then Exit ;
NatureG:=TOBL.GetValue('GL_NATUREPIECEG') ; Souche:=TOBL.GetValue('GL_SOUCHE') ;
Numero:=TOBL.GetValue('GL_NUMERO')        ; IndiceG:=TOBL.GetValue('GL_INDICEG') ;
NumL:=TOBL.GetValue('GL_NUMLIGNE')         ;
TOBLF.PutValue('GLF_NATUREPIECEG',NatureG) ; TOBLF.PutValue('GLF_SOUCHE',Souche) ;
TOBLF.PutValue('GLF_NUMERO',Numero)        ; TOBLF.PutValue('GLF_INDICEG',IndiceG) ;
TOBLF.PutValue('GLF_NUMLIGNE',NumL)   ;
for ind1:=0 to TOBLF.Detail.Count-1 do
  begin
  TOBDet:=TOBLF.Detail[ind1] ;
  TOBDet.PutValue('GLF_NATUREPIECEG',NatureG) ; TOBDet.PutValue('GLF_SOUCHE',Souche) ;
  TOBDet.PutValue('GLF_NUMERO',Numero)        ; TOBDet.PutValue('GLF_INDICEG',IndiceG) ;
  TOBDet.PutValue('GLF_NUMLIGNE',NumL)   ;
  end ;
end ;

procedure ValideLesFormules (TOBPiece,TOBLFormule : TOB);
Var ind,IndiceFormule : integer ;
    TOBLF,TOBL : TOB ;
begin
if TOBLFormule.Detail.Count<=0 then Exit ;
TOBLFormule.Detail[0].AddChampSup('UTILISE',True) ;
for ind:=0 to TOBPiece.Detail.Count-1 do
  begin
  TOBL:=TOBPiece.Detail[ind];
  if TOBL.FieldExists('GL_INDICEFORMULE') then
    IndiceFormule:=TOBL.GetValue('GL_INDICEFORMULE')
    else IndiceFormule:=0;
  if IndiceFormule>0 then
    begin
    TOBLF:=TOBLFormule.Detail[IndiceFormule-1] ;
    TOBLF.PutValue('UTILISE','X') ;
    UpdateLesFormules(TOBL,TOBLF) ;
    end ;
  end ;
for ind:=TOBLFormule.Detail.Count-1 downto 0 do
  begin
  TOBLF:=TOBLFormule.Detail[ind] ;
  if TOBLF.GetValue('UTILISE')<>'X' then TOBLF.Free ;
  end ;
if Not TOBLFormule.InsertDB(Nil) then V_PGI.IoError:=oeUnknown ;
end ;

function DetruitLesFormules (TOBPiece,TOBLFormule : TOB) : boolean;
Var St : string;
begin
Result:=True;
if TOBLFormule<>Nil then if TOBLFormule.Detail.Count>0 then
  begin
  St:=' GLF_NATUREPIECEG="'+TOBPiece.GetValue('GP_NATUREPIECEG')+'" AND GLF_SOUCHE="'
   + TOBPiece.GetValue('GP_SOUCHE')+'" '
   +' AND GLF_NUMERO='+IntToStr(TOBPiece.GetValue('GP_NUMERO'))+' AND GLF_INDICEG='
   +IntToStr(TOBPiece.GetValue('GP_INDICEG'))
   +' AND GLF_TYPEFORMULE="QTE" ' ;
  if ExecuteSQL('Delete FROM LIGNEFORMULE WHERE '+St)<=0 then Result:=False;
  end ;
end;

procedure ReliquatFormule (TOBLigne,TOBLFormule : TOB) ;
Var TOBLF,TOBDet : TOB;
    ind, IndiceFormule : integer ;
begin
if TOBLigne.FieldExists('GL_INDICEFORMULE') then
  IndiceFormule:=TOBLigne.GetValue('GL_INDICEFORMULE')
  else IndiceFormule:=0;
if IndiceFormule=0 then exit;
TOBLF:=TOBLFormule.Detail[IndiceFormule-1];
for ind:=0 to TOBLF.Detail.Count-1 do
  begin
  TOBDet:=TOBLF.Detail[Ind];
    if TOBDet.GetValue('GLF_ENABLEVAR')='X' then
       TOBDet.PutValue('GLF_VALEURVAR',0);
  end;
end;

end.
