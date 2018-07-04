unit factAcompte;

interface
uses Entgc,HEnt1,HCtrls,HPanel,UIUtil,Forms,ParamSoc,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  Doc_Parser,DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main,UserChg,
{$ENDIF}
  UTOB,Facture,FactUtil,uEntCommun
;
function GetMontantReliquatAcompte (TOBPiece : TOB) : Double;
function GetMontantAcompteInit (TOBPiece : TOB) : Double;
procedure AddLesChampSupAcptes (TOBN,TOBPiece : TOB);
Function SaisieAvancementAcompte ( FinTravaux,AcomptesObligatoire : boolean ;LesAcomptes : TOB;CleDoc : R_CleDoc ; Action : TActionFiche ; LeCodeTiers : String = '';Laffaire:string='';Lavenant:string='';SaisieAvanc: boolean =false; OrigineEXCEL: boolean =false) : boolean ;

implementation
uses UCumulCollectifs;

{***********A.G.L.***********************************************
Auteur  ...... : Lionel SANTUCCI

Créé le ...... : 04/10/2002

Modifié le ... :

Description .. : Appel d'une saisie d'avancement incluant la gestion des acomptes

Mots clefs ... : PIECE;SAISIE;

*****************************************************************}

Function SaisieAvancementAcompte ( FinTravaux,AcomptesObligatoire : boolean ;LesAcomptes : TOB;CleDoc : R_CleDoc ; Action : TActionFiche ; LeCodeTiers : String = '';Laffaire:string='';Lavenant:string='';SaisieAvanc: boolean =false; OrigineEXCEL: boolean =false) : boolean ;
var X : TFFacture ;
    PP  : THPanel ;
begin
  Result:=True ;
  SourisSablier;
  PP:=FindInsidePanel ;
  X:=TFFacture.Create(Application) ;

  X.LAffaire:=laffaire;
  X.lavenant:=lavenant;
  X.SaisieTypeAvanc := SaisieAvanc;
  X.OrigineEXCEL := OrigineEXCEL;
  X.ValidationSaisie:=False;
  X.LesAcomptes := LesAcomptes;
  X.AcompteObligatoire := AcomptesObligatoire;
  X.FinTravaux := FinTravaux;
  // ----
  X.CleDoc:=CleDoc ; X.Action:=Action ;
  X.NewNature:=X.CleDoc.NaturePiece ;
  X.LeCodeTiers:=LeCodeTiers ;
  X.TransfoPiece:=False ;
  X.DuplicPiece:=False ;

  if PP=Nil then
  BEGIN
    try
      X.ShowModal ;
    finally
      result:=X.ValidationSaisie;
      X.Free ;
    end ;
    SourisNormale ;
  END else
  BEGIN
    InitInside(X,PP) ;
    X.Show ;
  END ;
END;

procedure AddLesChampSupAcptes (TOBN,TOBPiece : TOB);
begin
TOBN.AddChampsupValeur ('NATURE',TOBPiece.GetValue('GP_NATUREPIECEG'),false);
TOBN.AddChampsupValeur ('NUMERO',TOBPiece.GetValue('GP_NUMERO'),false);
end;

function GetMontantReliquatAcompte (TOBPiece : TOB) : Double;
var QQ: TQuery;
begin
  result := 0;
  if TOBPiece.GetValue('GP_AFFAIREDEVIS') = '' then exit;
  TRY
    QQ := OpenSql('SELECT AFF_ACOMPTE,AFF_ACOMPTEREND FROM AFFAIRE WHERE AFF_AFFAIRE="' + TOBPiece.GetValue('GP_AFFAIREDEVIS') + '"', true,-1, '', True);
    if not QQ.eof then
    begin
      result := QQ.FindField('AFF_ACOMPTE').AsFloat - QQ.FindField('AFF_ACOMPTEREND').AsFloat;
      if Result < 0 then result := 0;
    end;
  FINALLY
    ferme(QQ);
  END;
end;

function GetMontantAcompteInit (TOBPiece : TOB) : Double;
var QQ: TQuery;
begin
  result := 0;
  if TOBPiece.GetValue('GP_AFFAIREDEVIS') = '' then exit;
  TRY
    QQ := OpenSql('SELECT AFF_ACOMPTE FROM AFFAIRE WHERE AFF_AFFAIRE="' + TOBPiece.GetValue('GP_AFFAIREDEVIS') + '"', true,-1, '', True);
    if not QQ.eof then
    begin
      result := QQ.FindField('AFF_ACOMPTE').AsFloat;
    end;
  FINALLY
    ferme(QQ);
  END;
end;

end.
