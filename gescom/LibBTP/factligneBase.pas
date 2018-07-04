unit factligneBase;

interface

uses HEnt1, UTOB,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HCtrls, FactTOB, SysUtils, EntGC ;

function AddMereLignesBases (NumOrdre : integer;TOBBasesL : TOB) : TOB;
procedure ChangeParentLignesBases (TOBBasesL,TOBtaxes : TOB; Nbdec : integer);
procedure DroplesBasesLigne(TOBL,TOBBasesL : TOB);
function FindLignesbases (NumOrdre: integer; TOBBasesL: TOB): TOB;
procedure OrdonnelignesBases (TOBBasesL : TOB);
procedure ValideLignesBases(TOBPiece,TOBmere : TOB);
procedure CumuleLigneBase (TOBDest,TOBSource : TOB);
function  GetTauxTaxeST (Fournisseur,FamilleTaxe : string) : double;
implementation

uses paramsoc,Ent1,FactCalc,ULiquidTva2014;
procedure CumuleLigneBase (TOBDest,TOBSource : TOB);
var indice : integer;
    TOBD,TOBS : TOB;
begin
  for Indice := 0 to TOBSource.detail.count -1 do
  begin
    TOBS := TOBSOurce.detail[indice];
    TOBD := TOBDest.findFirst(['BLB_CATEGORIETAXE','BLB_FAMILLETAXE'],
                              [TOBS.GetValue('BLB_CATEGORIETAXE'),TOBS.GetValue('BLB_FAMILLETAXE')],True);
    if TOBD = nil then
    begin
      TOBD := TOB.create ('LIGNEBASE',TOBDest,-1);
      TOBD.putvalue('BLB_CATEGORIETAXE',TOBS.GetValue('BLB_CATEGORIETAXE'));
      TOBD.putvalue('BLB_FAMILLETAXE',TOBS.GetValue('BLB_FAMILLETAXE'));
      TOBD.putvalue('BLB_TAUXTAXE',TOBS.GetValue('BLB_TAUXTAXE'));
      TOBD.putvalue('BLB_TAUXDEV',TOBS.GetValue('BLB_TAUXDEV'));
      TOBD.putvalue('BLB_COTATION',TOBS.GetValue('BLB_COTATION'));
      TOBD.putvalue('BLB_DEVISE',TOBS.GetValue('BLB_DEVISE'));
    end;
    TOBD.putValue('BLB_BASETAXE',TOBD.GetValue('BLB_BASETAXE')+TOBS.getValue('BLB_BASETAXE'));
    TOBD.putValue('BLB_BASEDEV',TOBD.GetValue('BLB_BASEDEV')+TOBS.getValue('BLB_BASEDEV '));
    TOBD.putValue('BLB_BASETAXETTC',TOBD.GetValue('BLB_BASETAXETTC')+TOBS.getValue('BLB_BASETAXETTC'));
    TOBD.putValue('BLB_BASETTCDEV',TOBD.GetValue('BLB_BASETTCDEV')+TOBS.getValue('BLB_BASETTCDEV '));
    TOBD.putValue('BLB_VALEURTAXE',TOBD.GetValue('BLB_VALEURTAXE')+TOBS.getValue('BLB_VALEURTAXE '));
    TOBD.putValue('BLB_VALEURDEV',TOBD.GetValue('BLB_VALEURDEV')+TOBS.getValue('BLB_VALEURDEV '));
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : L.S
Créé le ...... : 22/02/2008
Modifié le ... :   /  /
Description .. : Ajout d'une mere contenant les lignes de taxes rattachés à une ligne de piece
Mots clefs ... :
*****************************************************************}
function AddMereLignesBases (NumOrdre : integer;TOBBasesL : TOB) : TOB;
begin
  result := TOB.create('UNE MERE',TOBBasesL,-1);
  result.AddChampSupValeur ('NUMORDRE',NumOrdre);
end;



{***********A.G.L.***********************************************
Auteur  ...... : L.S
Créé le ...... : 22/02/2008
Modifié le ... :   /  /
Description .. : Cumul des taxes ligne sur les lignes
Mots clefs ... :
*****************************************************************}
procedure ChangeParentLignesBases (TOBBasesL,TOBtaxes : TOB; Nbdec : integer);
var Indice : integer;
    TOBT,TOBBL,TOBmereL : TOB;
    Categorie,Famille : string;
    Numordre : integer;
    Valeurcalc,Ecart,Taxe,TaxeST : double;
    IndiceMax : integer;
    Fournisseur,TypeInterv : string;
begin
  Indice := 0;
  if TOBtaxes.detail.count = 0 then exit;
  for Indice := 0 to TOBTaxes.detail.count -1 do
  begin
    TOBT := TOBTAxes.detail[Indice];
    //
    Fournisseur := TOBT.getValue('BLB_FOURNISSEUR');
    Numordre := TOBT.getValue('BLB_NUMORDRE');
    Categorie := TOBT.getValue('BLB_CATEGORIETAXE');
    Famille := TOBT.getValue('BLB_FAMILLETAXE');
    TypeInterv := TOBT.getValue('BLB_TYPEINTERV');
    //
    TOBmereL := FindLignesbases (NumOrdre,TOBBasesL);
    if TOBmereL = nil then TOBmereL := AddMereLignesBases (NumOrdre,TOBBasesL);
    //
    TOBBL := TOBmereL.findFIrst(['BLB_NUMORDRE','BLB_FOURNISSEUR','BLB_CATEGORIETAXE','BLB_FAMILLETAXE'],[Numordre,Fournisseur,Categorie,Famille],true);
    if TOBBL = nil then
    begin
      TOBBL:=TOB.create('LIGNEBASE',TOBMereL,-1);
      TOBBL.putValue('BLB_NUMORDRE',Numordre);
      TOBBL.putValue('BLB_FOURNISSEUR',Fournisseur);
      TOBBL.putValue('BLB_TYPEINTERV',TypeInterv);
      TOBBL.putValue('BLB_CATEGORIETAXE',Categorie);
      TOBBL.putValue('BLB_FAMILLETAXE',Famille);
      TOBBL.putValue('BLB_TAUXTAXE',TOBT.GetValue('BLB_TAUXTAXE'));
      TOBBL.putValue('BLB_TAUXDEV',TOBT.GetValue('BLB_TAUXDEV'));
      TOBBL.putValue('BLB_COTATION',TOBT.GetValue('BLB_COTATION'));
      TOBBL.putValue('BLB_DEVISE',TOBT.GetValue('BLB_DEVISE'));
    end;
    // Cas de la facturation TTC
    TOBBL.putValue('BLB_BASETAXETTC',TOBBL.GetValue('BLB_BASETAXETTC')+TOBT.GetValue('BLB_BASETAXETTC'));
    TOBBL.putValue('BLB_BASETTCDEV',TOBBL.GetValue('BLB_BASETTCDEV')+TOBT.GetValue('BLB_BASETTCDEV'));
    // --
    TOBBL.putValue('BLB_BASETAXE',TOBBL.GetValue('BLB_BASETAXE')+TOBT.GetValue('BLB_BASETAXE'));
    TOBBL.putValue('BLB_VALEURTAXE',TOBBL.GetValue('BLB_VALEURTAXE')+TOBT.GetValue('BLB_VALEURTAXE'));
    TOBBL.putValue('BLB_BASEDEV',TOBBL.GetValue('BLB_BASEDEV')+TOBT.GetValue('BLB_BASEDEV'));
    TOBBL.putValue('BLB_VALEURDEV',TOBBL.GetValue('BLB_VALEURDEV')+TOBT.GetValue('BLB_VALEURDEV'));
    //
    TOBBL.putValue('BLB_BASEACHAT',TOBBL.GetValue('BLB_BASEACHAT')+TOBT.GetValue('BLB_BASEACHAT'));
    TOBBL.putValue('BLB_VALEURACHAT',TOBBL.GetValue('BLB_VALEURACHAT')+TOBT.GetValue('BLB_VALEURACHAT'));
    //
    Taxe := (TOBBL.GetValue('BLB_TAUXTAXE')/100);
    //
    if TOBBL.getValue('BLB_BASETAXETTC')<>0 then
    begin
      // Cas de la facturation TTC --> On reequilibre la base Ht via le TTC
      ValeurCalc := arrondi(TOBBL.GetValue('BLB_BASETAXETTC')-(TOBBL.GetValue('BLB_BASETAXETTC')/(1+Taxe)),nbdec);
      Ecart := ValeurCalc-TOBBL.GetValue('BLB_VALEURTAXE');
      if Ecart <> 0 then
      begin
        TOBBL.putValue('BLB_VALEURTAXE',ValeurCalc);
      end;
      TOBBL.putValue('BLB_BASETAXE',TOBBL.GetValue('BLB_BASETAXETTC')-TOBBL.GetValue('BLB_VALEURTAXE'));
      //
      ValeurCalc := arrondi(TOBBL.GetValue('BLB_BASETTCDEV')-(TOBBL.GetValue('BLB_BASETTCDEV')/(1+Taxe)),nbdec);
      Ecart := ValeurCalc-TOBBL.GetValue('BLB_VALEURDEV');
      if Ecart <> 0 then
      begin
        TOBBL.putValue('BLB_VALEURDEV',ValeurCalc);
      end;
      TOBBL.putValue('BLB_BASEDEV',TOBBL.GetValue('BLB_BASETTCDEV')-TOBBL.GetValue('BLB_VALEURDEV'));
      //
    end else
    begin
      // Cas de la facturation HT ----> Equilibrage taxe ligne
      ValeurCalc := arrondi(TOBBL.GetValue('BLB_BASETAXE')*Taxe,nbdec);
      Ecart := ValeurCalc-TOBBL.GetValue('BLB_VALEURTAXE');
      if Ecart <> 0 then
      begin
        TOBBL.putValue('BLB_VALEURTAXE',ValeurCalc);
      end;
      ValeurCalc := arrondi(TOBBL.GetValue('BLB_BASEDEV')*Taxe,nbdec);
      Ecart := ValeurCalc-TOBBL.GetValue('BLB_VALEURTAXE');
      if Ecart <> 0 then
      begin
        TOBBL.putValue('BLB_VALEURDEV',ValeurCalc);
      end;
    end;
    
    if TypeInterv = 'Y00' then
    begin
//      if IsAutoLiquidationTvaST (Fournisseur) then TaxeST := 0
//      																			  else TaxeST := GetTauxTaxeST(Fournisseur,GetparamSocSecur('SO_BTTAXESOUSTRAIT','TN'));

      TaxeST := GetTauxTaxeST(Fournisseur,GetTvaST(Fournisseur));
    	ValeurCalc:=Arrondi(CalculeMontantTaxe(TOBBL.GetValue('BLB_BASEACHAT'),TaxeST,'',nil),NbDec);
      if IsAutoLiquidationTvaST (Fournisseur) then ValeurCalc := 0;
      TOBBL.putValue('BLB_VALEURACHAT',ValeurCalc);
    end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : L.S
Créé le ...... : 22/02/2008
Modifié le ... :   /  /
Description .. : Suppression des bases d'une ligne
Mots clefs ... :
*****************************************************************}
procedure DroplesBasesLigne(TOBL,TOBBasesL : TOB);
var TOBMere : TOB;
    TOBB : TOB;
begin
  TOBmere := FindLignesbases(TOBL.GEtValue('GL_NUMORDRE'),TOBBasesL);
  if TOBMere <> nil then TOBMere.clearDetail;
end;

{***********A.G.L.***********************************************
Auteur  ...... : L.S
Créé le ...... : 22/02/2008
Modifié le ... :   /  /
Description .. : Permet de retrouver les lignes de taxes rattachés à une ligne de piece
Mots clefs ... :
*****************************************************************}
function FindLignesbases (NumOrdre: integer; TOBBasesL: TOB) : TOB;
begin
	result := nil;
	if TOBBasesL = nil then exit;
  result := TOBBasesL.findFirst(['NUMORDRE'],[NumOrdre],false);
end;

{***********A.G.L.***********************************************
Auteur  ...... : L.S
Créé le ...... : 22/02/2008
Modifié le ... :   /  /
Description .. : Permet de remettre en forme les lignes de bases lues and la table
Mots clefs ... :
*****************************************************************}
procedure OrdonnelignesBases (TOBBasesL : TOB);
var Indice : integer;
    TOBL,TOBMere : TOB;
begin
  Indice := 0;
  if TOBBasesL.detail.count = 0 then exit;
  repeat
    TOBL := TOBBasesL.detail[indice];
    if TOBL.NomTable = 'LIGNEBASE' then
    begin
      TOBmere := FindLignesbases (TOBL.getValue('BLB_NUMORDRE'),TOBBasesL);
      if TOBmere = nil then
      begin
        TOBmere := AddMereLignesBases (TOBL.getValue('BLB_NUMORDRE'),TOBBasesL);
      end;
      TOBL.changeParent (TOBmere,-1);
    end else break;
  until Indice >= TOBBasesL.detail.count-1;
end;


{***********A.G.L.***********************************************
Auteur  ...... : L.S
Créé le ...... : 22/02/2008
Modifié le ... :   /  /
Description .. : Permet de rattacher les bases lignes au document
Mots clefs ... :
*****************************************************************}
procedure ValideLignesBases(TOBPiece,TOBmere : TOB);
var Indice : integer;
    TOBB : TOB;
begin
  for Indice := 0 to TOBmere.detail.count -1 do
  begin
    TOBB := TOBMere.detail[Indice];
    TOBB.PutValue('BLB_NATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG')) ;
    TOBB.PutValue('BLB_DATEPIECE',TOBPiece.GetValue('GP_DATEPIECE')) ;
    TOBB.PutValue('BLB_SOUCHE',TOBPiece.GetValue('GP_SOUCHE')) ;
    TOBB.PutValue('BLB_NUMERO',TOBPiece.GetValue('GP_NUMERO')) ;
    TOBB.PutValue('BLB_INDICEG',TOBPiece.GetValue('GP_INDICEG')) ;
  end;
end;

function  GetTauxTaxeST (Fournisseur,FamilleTaxe : string) : double;
var RegimeTaxe : string;
		TOBTa: TOB;
		QQ : TQuery;
begin
  Result := 0;
  if Fournisseur = '' then Exit;
  QQ := OpenSQL('SELECT T_REGIMETVA FROM TIERS WHERE T_TIERS="'+Fournisseur+'" AND T_NATUREAUXI="FOU"',True,1,'',true);
  if not QQ.eof then
  begin
    RegimeTaxe := QQ.FindField('T_REGIMETVA').AsString;
    TOBTa:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],['TX1',RegimeTaxe,FamilleTaxe],False) ;
    if TOBTa <> nil then result :=TOBTa.GetValue('TV_TAUXACH')
  end;
end;

end.
