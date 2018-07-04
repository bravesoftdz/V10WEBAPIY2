unit InvUtil;

interface
uses {$IFDEF VER150} variants,{$ENDIF} UTOB,uEntCommun;

type
  TTOBProc = function(T : TOB) : integer of object;
  TQtePrixRec = record
                  SomethingReturned : boolean;
                  DateDerniereCloture : TDateTime;
                  Qte,
                  QtePlus,
                  QteInv,
                  QteMoins : Double;
                  DPA,
                  PMAP,
                  DPR,
                  PMRP : Double;
                end;

function GetTOBValue(KelTOB : TOB; Field : String) : String;

// Si NumLot = '' -> article non géré par lot
// sinon artick géré par lot.
function GetStockOrdi(Article, Depot : String; NumLot : String = '') : Double;
function GetLastCloturedQtePrix(Article, Depot : String; DateListe : TDateTime) : TQtePrixRec; // var DateCloture : TDateTime) : TQtePrixRec;
function GetQtePrixDateListe(Article, Depot : String; DateListe : TDateTime ; TobL : TOB = nil ; TobLN : TOB = nil; TOBDispo : TOB = nil) : TQtePrixRec;
function GetPlusMoinsCpt(Nature, Cpt : String) : integer;
function GetPhyPlusMoins(Nature : String) : integer;

implementation
uses HCtrls,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     EntGC, HEnt1, StockUtil, UTofListeInv, SysUtils, UTofConsultStock, ParamSoc, FactComm,
  DateUtils;


{***********A.G.L.***********************************************
Auteur  ...... : TG
Créé le ...... : 18/02/2000
Modifié le ... :   /  /
Description .. : Récupère (en string) la valeur d'un champ d'une TOB avec retour vide si TOB nil
Mots clefs ... : CHAMP;TOB
*****************************************************************}
function GetTOBValue(KelTOB : TOB; Field : String) : String;
begin
if KelTOB = nil then result := ''
                else result := KelTOB.GetValue(Field);
end;


function GetStockOrdi(Article, Depot : String; NumLot : String = '') : Double;
var Q : TQuery;
begin
result := 0;
{MRNONGESTIONLOT}
NumLot:= '';
if NumLot = '' then
  begin
  Q := OpenSQL('SELECT GQ_PHYSIQUE AS QPHYSIQUE FROM DISPO '+
               'WHERE GQ_ARTICLE="'+Article+'" AND GQ_DEPOT="'+Depot+'"', true,-1, '', True);
//  if not Q.EOF then result := Q.FindField('GQ_PHYSIQUE').AsInteger;
  end else
  begin
  Q := OpenSQL('SELECT GQL_PHYSIQUE AS QPHYSIQUE FROM DISPOLOT '+
               'WHERE GQL_ARTICLE="'+Article+'" AND GQL_DEPOT="'+Depot+'" '+
               'AND GQL_NUMEROLOT="'+NumLot+'"', true,-1, '', True);
//  if not Q.EOF then result := Q.FindField('GQL_PHYSIQUE').AsInteger;
  end;
if not Q.EOF then result := Q.FindField('QPHYSIQUE').AsFloat;
Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Didier Carret
Créé le ...... : 20/03/2002
Modifié le ... : 20/03/2002
Description .. : Indique si la nature de pièce passée en paramètre affecte
Suite ........ : le compteur passé en paramètre en moins ou en plus
Mots clefs ... : STOCK;COMPTEUR;NATURE;PIECE
*****************************************************************}
function GetPlusMoinsCpt(Nature, Cpt : String) : integer;
var TOBNat : TOB ;
begin
result:=0 ;
TOBNat:=VH_GC.TOBParPiece.FindFirst(['GPP_NATUREPIECEG'],[Nature],False) ;
if TOBNat=nil then exit ;
if Pos(Cpt,TOBNat.GetValue('GPP_QTEPLUS')) > 0 then result:=1
else if Pos(Cpt,TOBNat.GetValue('GPP_QTEMOINS')) > 0 then result:=-1 ;
end;

// Retourne -1 si faut déstocker, 1 si faut stocker, 0 si rien
{***********A.G.L.***********************************************
Auteur  ...... : Guillaume THIBAUD
Créé le ...... : 11/07/2001
Modifié le ... : 11/07/2001
Description .. : Indique si une nature de pièce passée en paramètre affecte
Suite ........ : le stock physique en moins ou en plus
Mots clefs ... : STOCK;PHYSIQUE;NATURE;PIECE
*****************************************************************}
function GetPhyPlusMoins(Nature : String) : integer;
begin
result:=GetPlusMoinsCpt(Nature,'PHY') ;
end;

// Retourne Low(Double) si pas de cloture trouvée
function GetLastCloturedQtePrix(Article, Depot : String; DateListe : TDateTime) : TQtePrixRec;// var DateCloture : TDateTime)
var Q : TQuery;
    nbClot : integer;
begin

  result.SomethingReturned := false;
  //DateCloture := 0;

  //test si pas de clôture
  Q := OpenSQL('SELECT COUNT(GQ_ARTICLE) AS NBCLOT ' +
               'FROM DISPO WHERE GQ_ARTICLE="' + Article + '" ' +
               ' AND GQ_DEPOT="' + Depot + '" ' +
               ' AND GQ_CLOTURE="X"',true,-1, '', True);

  if not Q.EOF then
    nbClot := Q.FindField('NBCLOT').AsInteger
  else
    nbClot := 0;
  Ferme(Q);

//if nbClot = 0 then

//     Q := OpenSQL('SELECT GQ_DATECLOTURE, GQ_STOCKINITIAL AS QTESTOCK, GQ_DPA, GQ_PMAP, GQ_DPR, GQ_PMRP FROM DISPO '+
     Q := OpenSQL('SELECT GQ_DATEINV , GQ_STOCKINV, GQ_DPA, GQ_PMAP, GQ_DPR, GQ_PMRP FROM DISPO '+
             'WHERE GQ_ARTICLE="'+Article+'" AND GQ_DEPOT="'+Depot+'" AND GQ_CLOTURE="-"',true,-1, '', True);
(*
else Q := OpenSQL('SELECT GQ_DATECLOTURE AS DATECLOTURE, GQ_PHYSIQUE AS QTESTOCK, GQ_DPA, GQ_PMAP, GQ_DPR, GQ_PMRP FROM DISPO '+
             'WHERE GQ_ARTICLE="'+Article+'" AND GQ_DEPOT="'+Depot+'" AND GQ_CLOTURE="X" AND GQ_DATECLOTURE <= "'+USDateTime(DateListe)+'" '+
             'ORDER BY GQ_DATECLOTURE DESC', true,-1, '', True); // La 1ere réponse est la date la + récente
*)
if not Q.EOF then
  begin
  result.SomethingReturned := true;
  result.DateDerniereCloture := Q.FindField('GQ_DATEINV').AsDateTime;
  //DateCloture := Q.FindField('GQ_DATECLOTURE').AsDateTime;
  result.Qte := Q.FindField('GQ_STOCKINV').AsFloat;
  result.QtePlus := 0;
  result.QteMoins := 0;
  result.DPA := Q.FindField('GQ_DPA').AsFloat;
  result.PMAP := Q.FindField('GQ_PMAP').AsFloat;
  result.DPR := Q.FindField('GQ_DPR').AsFloat;
  result.PMRP := Q.FindField('GQ_PMRP').AsFloat;
  end;

  Ferme(Q);

end;

const Fieldz = 'GL_QUALIFQTESTO, GL_QUALIFQTEACH, GL_QUALIFQTEVTE, GL_NATUREPIECEG, '+
               'GL_DATEPIECE, GL_SOUCHE, GL_NUMERO, GL_INDICEG, GL_NUMLIGNE, GL_NUMORDRE,'+
               'GL_PIECEPRECEDENTE, GL_QTESTOCK, GL_QTERESTE, GL_MTRESTE, GL_QTERELIQUAT, GL_MTRELIQUAT, ' +
               'GL_DPA, GL_DPR,GL_PUHTNET, GL_PUTTCNET, GL_PRIXPOURQTE, GL_COEFCONVQTE ';

function GetParentQte(TOBLigneEtNomen : TOB) : Double;
var Q : TQuery;
    TOBParent : TOB;
begin
result := TOBLigneEtNomen.GetValue('GLN_QTE');
if (TOBLigneEtNomen.GetValue('GLN_COMPOSE') = null) or
   (TOBLigneEtNomen.GetValue('GLN_COMPOSE') = '') then exit;

Q := OpenSQL('SELECT GLN_NATUREPIECEG,GLN_SOUCHE,GLN_NUMERO, GLN_COMPOSE,GLN_INDICEG,GLN_NUMLIGNE, GLN_ORDRECOMPO, GLN_QTE '+
             'FROM LIGNENOMEN '+
             'WHERE GLN_ARTICLE="'+TOBLigneEtNomen.GetValue('GLN_COMPOSE')+'" '+
                   'AND GLN_NUMORDRE='+IntToStr(TOBLigneEtNomen.GetValue('GLN_ORDRECOMPO'))+' '+
                   'AND GLN_NATUREPIECEG="'+TOBLigneEtNomen.GetValue('GLN_NATUREPIECEG')+'" '+
                   'AND GLN_SOUCHE="'+TOBLigneEtNomen.GetValue('GLN_SOUCHE')+'" '+
                   'AND GLN_NUMERO='+IntToStr(TOBLigneEtNomen.GetValue('GLN_NUMERO'))+' '+
                   'AND GLN_INDICEG='+IntToStr(TOBLigneEtNomen.GetValue('GLN_INDICEG'))+' '+
                   'AND GLN_NUMLIGNE='+IntToStr(TOBLigneEtNomen.GetValue('GLN_NUMLIGNE')), true,-1, '', True);
TOBParent := TOB.Create('', nil, -1);
if not TOBParent.SelectDB('', Q) then begin TOBParent.Free; TOBParent := nil; end;
Ferme(Q);
if TOBParent <> nil then begin result := result * GetParentQte(TOBParent);
                               TOBParent.Free; end;
end;

//function SelectLigne(ContainerTOB : TOB; TokenizedCle,Article : String ; isNomen : boolean) : TOB;
function SelectLigne(ContainerTOB : TOB; CleDoc : R_CleDoc ; Article, Depot : String ; isNomen : boolean) : TOB;
var QLigne : TQuery;
    WhereLigne : string;
    ParentQte : double;
begin
  (*
  WhereLigne := 'GL_NATUREPIECEG=' + ReadTokenSt(TokenizedCle) + ' ' +
                  'AND GL_SOUCHE='  + ReadTokenSt(TokenizedCle) + ' ' +
                  'AND GL_NUMERO='  + ReadTokenSt(TokenizedCle) + ' ' +
                  'AND GL_INDICEG=' + ReadTokenSt(TokenizedCle) + ' ' +
                  'AND GL_NUMLIGNE=' + ReadTokenSt(TokenizedCle);
  *)

  WhereLigne := 'GL_NATUREPIECEG="' +  CleDoc.NaturePiece + '" ' +
                  'AND GL_SOUCHE="'  + CleDoc.Souche + '" ' +
                  'AND GL_NUMERO='  + inttostr(CleDoc.NumeroPiece) + ' ' +
                  'AND GL_INDICEG=' + inttostr(CleDoc.Indice) + ' ' +
//                  'AND GL_NUMLIGNE=' + inttostr(CleDoc.NumLigne) + ' ' +
                  'AND GL_NUMORDRE=' + inttostr(CleDoc.NumOrdre) + ' ' +
                  'AND GL_DEPOT="' + Depot +'"';

  if not isNomen then
  begin
    QLigne := OpenSQL('SELECT GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO, GL_INDICEG, ' +
           'GL_NUMLIGNE,GL_NUMORDRE,GL_QTESTOCK,GL_QTERESTE,GL_MTRESTE, GL_PIECEPRECEDENTE FROM LIGNE WHERE ' + WhereLigne, true,-1, '', True);

    if not QLigne.Eof then
    begin
      result := TOB.Create('LIGNE', ContainerTOB, -1);
      result.SelectDB('',QLigne);
    end else result := nil;
    Ferme(QLigne);
  end else
  begin
    QLigne := OpenSQL('SELECT GLN_NATUREPIECEG, GLN_SOUCHE,GLN_NUMERO,GLN_INDICEG,GLN_NUMLIGNE,GLN_ORDRECOMPO,'+
                 'GL_QTESTOCK,GL_QTERESTE,GL_MTRESTE,GL_PIECEPRECEDENTE,GLN_COMPOSE, GLN_QTE '+
                 'FROM LIGNENOMEN,LIGNE WHERE (GL_NATUREPIECEG=GLN_NATUREPIECEG AND GL_SOUCHE=GLN_SOUCHE AND GL_NUMERO=GLN_NUMERO AND GL_INDICEG=GLN_INDICEG AND GL_NUMLIGNE=GLN_NUMLIGNE) '+
                 'AND GLN_ARTICLE="'+Article+'" AND ' + WhereLigne  , true,-1, '', True);
    if not QLigne.Eof then
    begin
       result := TOB.Create('LIGNE', ContainerTOB, -1);
       result.SelectDB('',QLigne);
       result.PutValue('GL_QTESTOCK', 0);
       result.PutValue('GL_QTERESTE', 0);
       result.PutValue('GL_MTRESTE', 0);
    end else result := nil;
    Ferme(QLigne);
    if result <> nil then
    begin
      ParentQte := GetParentQte(result);
      result.PutValue('GL_QTESTOCK', Result.GetValue('GL_QTESTOCK') * ParentQte);
      result.PutValue('GL_QTERESTE', Result.GetValue('GL_QTERESTE') * ParentQte);
    end;
  end;
end;

function GetQtePrixDateListe(Article, Depot : String; DateListe : TDateTime ; TobL : TOB = nil ; TobLN : TOB = nil; TOBDispo : TOB = nil) : TQtePrixRec;
var Q : TQuery;
    TOBBefore,TobSelect,TobNomen,TOBLigne,TOBNew, TOBLignePrec : TOB;
    itob, SensPiece : integer;
    FUS,FUA,FUV,Ratio,CoefUaUs : Double;
    ParentQte, Qte, QtePrec, PA, PR,PrixPourQte : Double;
    Cle, stWhere, VenteAchat,Requete : String;
    CleDocPrec : R_CleDoc;
    yy,mm,dd,hh,mn,ss,ms : word;
    DateDebutTrait : TDateTime;
begin
if TOBDispo <> nil then
begin
  Result.DateDerniereCloture := TOBDispo.GetDateTime('GQ_DATEINV');
  result.SomethingReturned := true;
  result.Qte := TOBDispo.GetDouble('GQ_STOCKINV');
  result.QtePlus := 0;
  result.QteMoins := 0;
  result.DPA := TOBDispo.GetDouble('GQ_DPA');
  result.PMAP := TOBDispo.GetDouble('GQ_PMAP');
  result.DPR := TOBDispo.GetDouble('GQ_DPR');
  result.PMRP := TOBDispo.GetDouble('GQ_PMRP');
end else
begin
	result := GetLastCloturedQtePrix(Article, Depot, DateListe);//, DC);
end;

//if not result.SomethingReturned then DC := 0;
if (DateListe = result.DateDerniereCloture ) then exit;

DecodeDateTime (result.DateDerniereCloture,yy,mm,dd,hh,mn,ss,ms);
DateDebutTrait := EncodeDateTime(yy,mm,dd,0,0,0,0);

TOBSelect := TOB.Create('prout', nil, -1);
if TobL = nil then
   begin
   stWhere := RecupWhereNaturePiece('PHY','GL',True);
   if stWhere <> '' then stWhere := ' AND ' +stWhere;
   Requete := 'SELECT '+Fieldz+' FROM LIGNE  '+
             'WHERE GL_ARTICLE="'+Article+'" AND GL_DEPOT="'+Depot+'" '+
                   'AND GL_TENUESTOCK="X" '+
                   'AND GL_QUALIFMVT<>"ANN" '+
                   'AND GL_DATEPIECE>="'+USDATETime(DateDebutTrait)+'" '+
                   'AND GL_DATEPIECE<="'+USDateTime(DateListe)+'" '+ stWhere +
             'ORDER BY GL_DATEPIECE';
   Q := OpenSQL(requete, true,-1, '', True);
   TOBSelect.LoadDetailDB('LIGNE', '', '', Q, false, false);
   Ferme(Q);
   end else TobSelect.Dupliquer(TobL,true,true);

//------------------------------------------------------------------------------
TOBNomen := TOB.Create('prout', nil, -1);
if TOBLN = nil then
   begin
   stWhere := RecupWhereNaturePiece('PHY','GLN',True);
   if stWhere <> '' then stWhere := ' AND ' +stWhere;
   Q := OpenSQL('SELECT '+Fieldz+', GLN_QUALIFQTESTO , GLN_QUALIFQTEACH , GLN_QUALIFQTEVTE , GLN_COMPOSE, GLN_NATUREPIECEG, GLN_SOUCHE,GLN_NUMERO,GLN_INDICEG,GLN_NUMLIGNE,GLN_ORDRECOMPO, GLN_QTE '+
               'FROM LIGNENOMEN,LIGNE WHERE (GL_NATUREPIECEG=GLN_NATUREPIECEG AND GL_SOUCHE=GLN_SOUCHE AND GL_NUMERO=GLN_NUMERO AND GL_INDICEG=GLN_INDICEG AND GL_NUMLIGNE=GLN_NUMLIGNE) '+
                   'AND GLN_ARTICLE="'+Article+'" AND GL_DEPOT="'+Depot+'" '+
                   'AND GLN_TENUESTOCK="X" '+
                   'AND GL_QUALIFMVT<>"ANN" '+
                   'AND GL_DATEPIECE>"'+USDateTime(DateDebutTrait)+'" '+
                   'AND GL_DATEPIECE<="'+USDateTime(DateListe)+'" '+ stWhere +
             ' ORDER BY GL_DATEPIECE', true,-1, '', True);
{   Q := OpenSQL('SELECT '+Fieldz+', GLN_QUALIFQTESTO , GLN_QUALIFQTEACH , GLN_QUALIFQTEVTE , GLN_COMPOSE, GLN_NATUREPIECEG, GLN_SOUCHE,GLN_NUMERO,GLN_INDICEG,GLN_NUMLIGNE,GLN_ORDRECOMPO, GLN_QTE '+
               'FROM LIGNE LEFT JOIN LIGNENOMEN ON (GL_NATUREPIECEG=GLN_NATUREPIECEG AND GL_SOUCHE=GLN_SOUCHE AND GL_NUMERO=GLN_NUMERO AND GL_INDICEG=GLN_INDICEG AND GL_NUMLIGNE=GLN_NUMLIGNE) '+
             'WHERE GLN_ARTICLE="'+Article+'" AND GL_DEPOT="'+Depot+'" '+
                   'AND GLN_TENUESTOCK="X" '+
                   'AND GL_QUALIFMVT<>"ANN" '+
                   'AND GL_DATEPIECE>"'+USDateTime(result.DateDerniereCloture)+'" '+
                   'AND GL_DATEPIECE<="'+USDateTime(DateListe)+'" '+ stWhere +
             'ORDER BY GL_DATEPIECE', true);  }
   TOBNomen.LoadDetailDB('LIGNE', '', '', Q, false, false);
   Ferme(Q);
   end else TOBNomen.Dupliquer(TobLN,true,true);

for itob := 0 to TOBNomen.Detail.Count-1 do
  begin
  TOBLigne := TOBNomen.Detail[itob];
  with TOBLigne do
   TOBNew := TOBSelect.FindFirst(['GL_NATUREPIECEG', 'GL_DATEPIECE', 'GL_SOUCHE', 'GL_NUMERO', 'GL_INDICEG', 'GL_NUMLIGNE'],
                                 [GetValue('GL_NATUREPIECEG'), GetValue('GL_DATEPIECE'), GetValue('GL_SOUCHE'), GetValue('GL_NUMERO'), GetValue('GL_INDICEG'), GetValue('GL_NUMLIGNE')],
                                 false);
  if TOBNew = nil then
    begin
    TOBNew := TOB.Create('LIGNE', TOBSelect, -1);
    TOBNew.Dupliquer(TOBLigne, true, true);
    TOBNew.PutValue('GL_QTESTOCK', 0);
    TOBNew.PutValue('GL_QTERESTE', 0);
    TOBNew.PutValue('GL_MTRESTE', 0);
    TOBNew.PutValue('GL_QTERELIQUAT', 0);
    TOBNew.PutValue('GL_MTRELIQUAT', 0);
    // Question de la qte reliquat ??
    end;

  ParentQte := GetParentQte(TOBLigne);
  TOBNew.PutValue('GL_QTESTOCK', TOBNew.GetValue('GL_QTESTOCK') +
                                 (TOBLigne.GetValue('GL_QTESTOCK') * ParentQte) );
  TOBNew.PutValue('GL_QTERESTE', TOBNew.GetValue('GL_QTERESTE') +
                                 (TOBLigne.GetValue('GL_QTERESTE') * ParentQte) );
  TOBNew.PutValue('GL_QTERELIQUAT', TOBNew.GetValue('GL_QTERELIQUAT') +
                                 (TOBLigne.GetValue('GL_QTERELIQUAT') * ParentQte) );

  TOBNew.AddChampSup('COMPOSANT',False);
  end;

TOBSelect.Detail.Sort('GL_DATEPIECE');
TOBNomen.ClearDetail;
TOBNomen.Free;
//------------------------------------------------------------------------------

TOBBefore := TOB.Create('prout', nil, -1);
for itob := 0 to TOBSelect.Detail.Count-1 do
  begin
  TOBLigne := TOBSelect.Detail[itob];
//  if (TOBLigne.GetValue('GL_DATEPIECE') > DC) {and (TOBLigne.GetValue('GL_DATEPIECE') <= DateListe)} then
    begin
    SensPiece := GetPhyPlusMoins(TOBLigne.GetValue('GL_NATUREPIECEG'));
    if SensPiece = 0 then continue;

    Cle := VarToStr(TOBLigne.GetValue('GL_PIECEPRECEDENTE'));  // MODIF LM ORACLE
    if Cle <> '' then DecodeRefPiece(Cle, CleDocPrec);

    Qte := TOBLigne.GetValue('GL_QTESTOCK');
    if (TOBLigne.GetValue('GL_INDICEG') = 0) and (Cle <> '') and (GetPhyPlusMoins(CleDocPrec.NaturePiece) <> 0) then
    begin
      TOBLignePrec := TOBSelect.FindFirst(['GL_NATUREPIECEG', 'GL_SOUCHE', 'GL_NUMERO', 'GL_INDICEG', 'GL_NUMORDRE'],
         [CleDocPrec.NaturePiece, CleDocPrec.Souche, intToStr(CleDocPrec.NumeroPiece), intToStr(CleDocPrec.Indice), intToStr(CleDocPrec.NumOrdre)], false);
      if TOBLignePrec = nil then // Pièce précédente avant date cloture?
         TOBLignePrec := SelectLigne(TOBBefore, CleDocPrec, Article, Depot, TOBSelect.Detail[itob].FieldExists('COMPOSANT')); // Faire un select exprès pour
      if TOBLignePrec <> nil then
      begin
        if not TOBLignePrec.FieldExists ('QTETRANS') then TobLignePrec.AddChampSupValeur ('QTETRANS', 0); {DBR NEWPIECE}
        QtePrec := TOBLignePrec.GetValue('GL_QTESTOCK');
        // Ce test est nécessaire si on modifie la qté livrée après avoir facturé
        // dans le cas ou on gère les reliquats sur la facture par exemple
{DBR NEWPIECE DEBUT}
        if Qte > QtePrec - TobLignePrec.GetValue ('QTETRANS') then
//        if Qte > QtePrec then
        begin
//          Qte := SensPiece * (Qte - QtePrec);
          Qte := SensPiece * (Qte - (QtePrec - TobLignePrec.GetValue ('QTETRANS')));
        end
{DBR NEWPIECE FIN}
        else Qte := 0;  // On ne peut sortir moins que ce qui a déjà été livré (pièce précédente)
        TobLignePrec.PutValue ('QTETRANS', TobLignePrec.GetValue ('QTETRANS') + TOBLigne.GetValue('GL_QTESTOCK')); {DBR NEWPIECE}
      end;
    end else Qte := (SensPiece * Qte);

    if Qte = 0 then Continue;

    TOBLigne := TOBSelect.Detail[itob];

    FUS := RatioMesure('PIE', TOBLigne.GetValue('GL_QUALIFQTESTO'));   // MODIF LM ORACLE
    FUV := RatioMesure('PIE', TOBLigne.GetValue('GL_QUALIFQTEVTE'));   // MODIF LM ORACLE
    FUA := RatioMesure('PIE', TOBLigne.GetValue('GL_QUALIFQTEACH'));   // MODIF LM ORACLE
    if TobLigne.FieldExists('GLN_QUALIFQTESTO') then
         Ratio:=GetRatio (TOBLigne,TOBLigne, trsStock)    //nomenclature
    else Ratio:=GetRatio (TOBLigne,nil, trsStock);

    VenteAchat:=GetInfoParPiece(TOBLigne.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT') ;      // MODIF LM ORACLE

    if VenteAchat = 'ACH' then
        begin
        CoefuaUs := TOBLigne.GetValue('GL_COEFCONVQTE');
        Qte:=Qte/Ratio;
        result.QtePlus := result.QtePlus + Qte;
        //result.DPA := Ratioize(TOBLigne.GetValue('GL_DPA'), FUS, FUA, 1.0);
        //Modif pour avoir le PUTTC ou PUHT (selon paramètrage) à la pièce
        PrixPourQte:= TOBLigne.GetValue('GL_PRIXPOURQTE') ;
        if PrixPourQte=0 then PrixPourQte:=1;
        if GetParamSoc('SO_GCACHATTTC') then
           result.DPA := Ratioize(TOBLigne.GetValue('GL_PUTTCNET')/PrixPourQte, FUS, FUA, 1.0)
           else result.DPA := Ratioize(TOBLigne.GetValue('GL_PUHTNET')/PrixPourQte, FUS, FUA, 1.0);
        result.DPR := Ratioize(TOBLigne.GetValue('GL_DPR'), FUS, FUA, 1.0);
        PA := result.DPA;
        PR := result.DPR;
        if (result.Qte < 0.0000001) and (result.Qte > -0.0000001) then Result.Qte := 0.0;
        if result.Qte <= 0 then
          begin
          result.PMAP := PA;
          result.PMRP := PR;
          end else
          if (result.Qte + Qte) > 0 then
            begin
            result.PMAP := ((result.PMAP * result.Qte) + (PA * Qte)) / (result.Qte + Qte);
            result.PMRP := ((result.PMRP * result.Qte) + (PR * Qte)) / (result.Qte + Qte);
            end;
        end
     else if VenteAchat = 'VEN'
            then begin
                 Qte:=Qte/Ratio;
                 result.QteMoins := result.QteMoins - Qte;
                 end;

    result.Qte := result.Qte + Qte;
    result.SomethingReturned := true;
    end;
  end;
if (result.Qte < 0.0000001) and (result.Qte > -0.0000001) then Result.Qte := 0.0;
if (result.QtePlus < 0.0000001) and (result.QtePlus > -0.0000001) then Result.QtePlus := 0.0;
if (result.QteMoins < 0.0000001) and (result.QteMoins > -0.0000001) then Result.QteMoins := 0.0;
Result.DPA := Arrondi(Result.DPA,V_PGI.OkDecV+2);
Result.PMAP := Arrondi(Result.PMAP,V_PGI.OkDecV+2);
Result.DPR := Arrondi(Result.DPR,V_PGI.OkDecV+2);
Result.PMRP := Arrondi(Result.PMRP,V_PGI.OkDecV+2);
TOBSelect.ClearDetail;
TOBSelect.Free;
TOBBefore.ClearDetail;
TOBBefore.Free;
end;

end.
