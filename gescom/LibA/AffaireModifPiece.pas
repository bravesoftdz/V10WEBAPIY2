unit AffaireModifPiece;

interface

uses HEnt1,FactComm,Dicobtp,FactGrp,FactNomen,FactCpta,FactCalc,FactPiece,
     AffaireUtil,Hctrls,SAISUTIL,Sysutils,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     UTob ,TarifUtil,Ent1,UtilPgi,FactUtil, EntGC, UtilGrp, FactTOB,uEntCommun,UtilTOBPiece ;

Type T_TypeModifPieceAff =  (tmpAugment,tmpAjoutLigne,tmpRecalcul,tmpPassageEuro);

Function  ModifieLaPieceAff (TypeModif : T_TypeModifPieceAff; CodeAff : string; TOBPieceRef : TOB; bParle: Boolean; Var Arg : string) : integer;

// Fonctions de conversion en Euro
Procedure ConvertEuro ( TobF : TOB ; NomT : String; bTenueEuro:boolean );
Procedure BourreLeDEV ( TobF : TOB ; NomChamp : String ; bParLeCon:boolean) ;
Procedure BourreLeChamp ( TobF : TOB ; NomChampOld, NomChampNew : String ) ;
Procedure SwapLeFlag ( TobF : TOB ; NomChamp : String ) ;
Function  CalcPourAugm(pour,mm : double;Annul:Boolean) : double;

Type T_ModifPieceAff = Class
       Constructor Create(TypeModif : T_TypeModifPieceAff; CodeAff,Arg : string; PieceRef : TOB; bParle : Boolean);
       private
       TypeM : T_TypeModifPieceAff;
       Aff, RecupTotaux : string;
       Parle : Boolean;
       NumError : integer;
       CD : R_CleDoc ;
       DEV    : RDEVISE ;
       Pour :double;
       Annul:boolean;
       Zarrondi : string;
       TobPieceRef,
       TOBPiece,TOBArticles,TOBCataLogu,TOBBases,TOBBasesL,TOBTiers,TOBEches,TOBAcomptes, TOBOuvrage, TOBOuvragesP,
       TOBPorcs,TOBNomenclature,TOBComms,TOBCpta,TOBAnaP,TOBAnaS,TOBAdresses,TOBTarif,TOBPieceTrait,TOBSSTRAIT : TOB;
       TobLigneTarif: Tob;
       procedure CalculPort ;
       Function  ModifPiece : integer;
       Procedure  ValideModifPiece;
       Procedure ToutAllouer;
       Procedure ToutLiberer;
       Function  ChargeTobPiece : Boolean;
       Function  AjouteLignes : Boolean;
       Function  RecopieLignes : Boolean;
       Function  AugmentePiece : Boolean;
       Function  PassageEuro(AcbTenueEuro:boolean) : Boolean;
       Procedure ChargeTobLignes (TobP : TOB; CleDoc : R_CleDoc);
       Procedure MajCleLigne (TOBL : TOB);
       Procedure InitToutModif ;
       Procedure MajLesComms ( TOBPiece : TOB ) ;
       end;
Const	MsgCreat: array[1..5] of string 	= (
          {1}        'La modification ne s''est pas complètement effectuée',
          {2}        'La pièce n''a  pas pu être enregistrées',
          {3}        'ATTENTION :',
          {4}        'Validation de l''affaire impossible',
          {5}        'La pièce associée à l''affaire n''est pas valide'
          );



implementation
uses factouvrage,FactureBTP,UCotraitance;

Function  ModifieLaPieceAff (TypeModif : T_TypeModifPieceAff; CodeAff : string; TOBPieceRef : TOB; bParle: Boolean; Var Arg : string) : integer;
Var ModifPieceAff : T_ModifPieceAff;
begin
Result := 0;
if (TypeModif = tmpAjoutLigne) and (TOBPieceRef = Nil) then Exit;
if CodeAff = '' then Exit;

ModifPieceAff := T_ModifPieceAff.Create(TypeModif,CodeAff,Arg,TobPieceRef,bParle);
Result := ModifPieceAff.ModifPiece;
Arg := ModifPieceAff.RecupTotaux;
ModifPieceAff.Free ;
end;


Constructor T_ModifPieceAff.Create(TypeModif : T_TypeModifPieceAff; CodeAff,Arg : string; PieceRef : TOB; bParle : Boolean);
var   Tmp,Champ, Valeur : string;
    X : integer;
begin
// Initialisation
NumError := 0;
// Recup Arguments
TypeM := TypeModif;
Aff   := CodeAff;
TobPieceRef := PieceRef;
Parle := bParle;
Tmp:=(Trim(ReadTokenSt(Arg)));
While (Tmp <> '') do
    begin
    X:=pos(':',Tmp);
    if x<>0 then
    	 begin
        Champ:=copy(Tmp,1,X-1);
        Valeur:=Copy (Tmp,X+1,length(Tmp)-X);
       end
    else
    	 Champ := Tmp;
    // Recup des arguments ...
    if Champ = 'POUR' then Pour:= StrToFloat(Valeur);
    if Champ = 'ARRONDI'    then Zarrondi:= Valeur;
    if Champ = 'ANNUL'    then If Valeur = 'TRUE' then  Annul:= True else Annul:=False;
    Tmp:=(Trim(ReadTokenSt(Arg)));
    end;
end;


Function T_ModifPieceAff.ModifPiece : Integer;
Var   io   : TIoErr ;
      bTraiteOK : Boolean;
   DateDevise : TDateTime;
begin
ToutAllouer;
bTraiteOk := True;
io := oeOK;

if bTraiteOk = True then bTraiteOk := ChargeTobPiece; // complément de la pièce
{Traitement des devises}
DEV.Code:=TOBPiece.GetValue('GP_DEVISE') ;
GetInfosDevise(DEV) ;
DateDevise :=  TobPiece.GetValue ('GP_DATEPIECE');
DEV.Taux:= GetTaux(DEV.Code,DEV.DateTaux,DateDevise) ;
// Ajout de ligne de la pièce de référence
if (TypeM = tmpAjoutLigne) then
   begin
   if bTraiteOk = True then bTraiteOk := AjouteLignes;
   end
else
// Augmentation de l'affaire
if (TypeM = tmpAugment) then
   begin
   if bTraiteOk = True then bTraiteOk := AugmentePiece;
   end
else
// Passage en Euro des affaires
if (TypeM = tmpPassageEuro) then
   begin
   if bTraiteOk = True then bTraiteOk := PassageEuro(VH^.TenueEuro);
   end;

if bTraiteOk = True then io:=Transactions(ValideModifPiece,1);
Case io of
   oeUnknown : Numerror:=1;
   oeSaisie :  Numerror:=2;
   END ;
if Parle then
   BEGIN
    if NumError > 0 then PgiBoxAf(MsgCreat[3]+ ': '+MsgCreat[NumError],MsgCreat[4]);
   END;

if NumError = 0 then
   begin
   RecupTotaux := 'GP_TOTALHTDEV:'+ FloattoStr(TobPiece.GetValue('GP_TOTALHTDEV')) +
                  ';GP_TOTALTTCDEV:'+ FloatToStr(TobPiece.GetValue('GP_TOTALTTCDEV'));
   end;
ToutLiberer;
Result := NumError;
end;
// *************************** Chargement **************************************
Procedure T_ModifPieceAff.ToutAllouer ;
begin      // voir fct ToutAllouer de FActure  ????
TOBPiece:=TOB.Create('PIECE',Nil,-1) ;
TOBArticles:=TOB.Create('',Nil,-1) ;
TOBCataLogu:=TOB.Create('',Nil,-1) ;
// mcd 28/01/03 TOBBases:=TOB.Create('BASES',Nil,-1) ;
TOBBases:=TOB.Create('les bases',Nil,-1) ;
TOBBasesL:=TOB.Create('les bases lignes',Nil,-1) ;
TOBTiers:=TOB.Create('TIERS',Nil,-1) ; TOBTiers.AddChampSup('RIB',False) ;
TOBEches:=TOB.Create('les echeances',Nil,-1) ;
TOBAcomptes:=TOB.Create('',Nil,-1) ;
TOBPorcs:=TOB.Create('',Nil,-1) ;
// mcd 28/01/02TOBNomenclature:=TOB.Create('NOMENCLATURES',Nil,-1) ;
TOBNomenclature:=TOB.Create('Les Nomenclatures',Nil,-1) ;
TOBComms:=TOB.Create('COMMERCIAUX',Nil,-1) ;
TOBCpta:=TOB.Create('',Nil,-1) ;
TOBAnaP:=TOB.Create('',Nil,-1) ;
TOBAnas:=TOB.Create('',Nil,-1) ;
TOBAdresses:=TOB.Create('',Nil,-1);
TOBTarif:=TOB.Create('TARIF',Nil,-1) ;
TOBLigneTarif:=TOB.Create('LIGNETARIF',Nil,-1) ;
TOBOUvrage := TOB.Create('LES OUVRAGES',nil,-1);
TOBOUvragesP := TOB.Create('LES OUVRAGES PLAT',nil,-1);
TOBPieceTrait := TOB.Create ('LES PIECES EXTE',nil,-1);
TOBSSTrait := TOB.Create ('LES SOUS TRAITANTS',nil,-1);
end;

Procedure T_ModifPieceAff.ToutLiberer ;
BEGIN
TOBOuvrage.free; TOBOuvragesP.free;
TOBPiece.Free ; TOBArticles.Free ;
TOBBases.Free ; TOBBasesL.free; TOBTiers.Free ; TOBNomenclature.Free ;
TOBAdresses.Free ; TOBCataLogu.Free;
TOBCpta.Free ; TOBAnaP.Free ; TOBAnaS.Free ; TOBEches.Free ;
TOBAcomptes.Free ; TOBPorcs.Free ; TOBComms.Free ;
TOBTarif.Free ;
TOBLigneTarif.Free ;
TOBPieceTrait.free;
TOBSSTrait.free;
END ;

Procedure T_ModifPieceAff.InitToutModif ;
Var NowFutur : TDateTime ;
begin    // voir IntiToutModif Facture.pas
NowFutur:=NowH ;  //mcd 22/11/02 idem facture
TOBPiece.SetAllModifie(True)    ; TOBPiece.SetDateModif(NowFutur) ;
if TOBOuvrage <> nil then TOBOuvrage.setAllModifie(true);
if TOBOuvragesP <> nil then TOBOuvragesP.setAllModifie(true);
TOBAdresses.SetAllModifie(True) ;
TOBBases.SetAllModifie(True)    ;
TOBBasesL.setAllModifie (true);
TOBEches.SetAllModifie(True)    ;
TOBAcomptes.SetAllModifie(True) ;
TOBPorcs.SetAllModifie(True)    ;
TOBAnaP.SetAllModifie(True)     ;
TOBAnaS.SetAllModifie(True)     ;
TOBNomenclature.SetAllModifie(True)  ;
TOBPieceTrait.SetAllModifie(True);
TOBSSTRAIT.SetAllModifie(True);
end;


Function T_ModifPieceAff.ChargeTobPiece : Boolean;
Var Q : TQuery ;
    i  : integer ;
    TOBL : TOB ;
    statut,typecom : string;
    IndiceOuv : integer;
BEGIN    // voir fct LoadLesTob de Facture.pas  ???
Result := False;
// chargement de la piece
if Copy(Aff,1,1) = 'P' then
	 Statut:= 'PRO'
Else if Copy(Aff,1,1) = 'I' then
	 Statut:= 'INT'
else
	 Statut := 'AFF';

if Not(GetTobPieceAffaire(Aff, Statut,TobPiece)) then
  begin
  	  Numerror:=5;
      Exit;
  end;
// Alim clédoc
CD:=TOB2CleDoc(TOBPiece);

// Chargement des lignes
ChargeTobLignes (TobPiece, CD);
// Articles
UG_AjouteLesArticles(TOBPiece,TOBArticles,TOBCpta,TOBTiers,TOBAnaP,TOBAnas,TOBCatalogu,False) ;
{$IFNDEF RECALCULAFF}
// Comms
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ;
    {commercial}  //gm
		if TypeCom='' then if ctxFO in V_PGI.PGIContexte then TypeCom:='VEN' else TypeCom:='REP' ;
    //gm
    AjouteRepres(TOBL.GetValue('GL_REPRESENTANT'),TypeCom,TOBComms) ;
    END ;
// Acomptes
LoadLesAcomptes (TOBPiece,TOBAcomptes,CD);
// Porcs   ** modif, il faut tout prendre ...
Q:=OpenSQL('SELECT * FROM PIEDPORT WHERE '+WherePiece(CD,ttdPorc,False),True,-1,'',true) ;
TOBPorcs.LoadDetailDB('PIEDPORT','','',Q,False) ;
Ferme(Q) ;
{$ENDIF}
// Eche  ** modif, il faut tout prendre ...
Q:=OpenSQL('SELECT * FROM PIEDECHE WHERE '+WherePiece(CD,ttdEche,False),True,-1,'',true) ;
TOBEches.LoadDetailDB('PIEDECHE','','',Q,False) ;
Ferme(Q) ;
// Adresses
// LoadLesAdresses(TOBPiece,TOBAdresses) ;
// Nomenclatures
LoadLesNomen(TOBPiece,TOBNomenclature,TOBArticles,CD) ;
// Lecture Analytiques
//LoadLesAna(TOBPiece,TOBAnaP,TOBAnaS) ;
IndiceOuv := 1;
LoadLesOuvrages(TOBPiece, TOBOuvrage, TOBArticles, CD, IndiceOuv);
LoadLaTOBPieceTrait(TOBpieceTrait,CD, '');
Result := True;
END;

Procedure T_ModifPieceAff.ChargeTobLignes (TobP : TOB; CleDoc : R_CleDoc);
Var Q : TQuery ;
    //i : integer;
    //TOBL : TOB;
begin     // voir LoadLesTob de Facture.pas ???
  //** modif, il faut tout prendre ..
Q:=OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',True,-1,'',true) ;
if Not Q.EOF then TOBP.LoadDetailDB('LIGNE','','',Q,True,True) ;
Ferme(Q) ;

PieceAjouteSousDetail(TOBP);   //gm 08/02/02

END;


//******************** Traitement ajout de lignes ******************************
Function T_ModifPieceAff.AjouteLignes : Boolean;
Var CDRef : R_CleDoc;
//    i : integer;
//    TobL : TOB;
BEGIN
Result := True;
CDRef:=TOB2CleDoc(TOBPieceRef);
// Chargement des lignes à intégrer
ChargeTobLignes (TobPieceRef, CDRef);
// Articles
UG_AjouteLesArticles(TOBPieceRef,TOBArticles,TOBCpta,TOBTiers,TOBAnaP,TOBAnaS,TOBCatalogu,False);
// Nomenclatures
LoadLesNomen(TOBPiece,TOBNomenclature,TOBArticles,CDRef) ;
Recopielignes;
// Renumérotation des lignes
NumeroteLignesGC (Nil, TobPiece);
END;

Function T_ModifPieceAff.RecopieLignes : Boolean;
Var   i : integer;
      TOBL,TOBLNew : TOB;
BEGIN
Result := True;
For i:=0 To TobPieceRef.Detail.Count -1 do
   begin
   TOBL:=TobPieceRef.Detail[i] ;
   TOBLNew := TOB.Create('LIGNE',TOBPiece,-1) ;
   TOBLNew.Dupliquer (TOBL,True,True);
   MajCleLigne (TOBLNew);
   end;
// Voir les nomenc ???
END;



Procedure T_ModifPieceAff.MajCleLigne (TOBL : TOB);
Var RefPiece : string;
begin
PieceVersLigne (TOBPiece,TobL);
{Traçabilité}
RefPiece := EncodeRefPiece(TobPieceRef); TOBL.PutValue('GL_PIECEPRECEDENTE',RefPiece) ;
if TOBL.GetValue('GL_PIECEORIGINE')='' then TOBL.PutValue('GL_PIECEORIGINE',RefPiece) ;

end;


Function T_ModifPieceAff.AugmentePiece : Boolean;
var     TobDet : TOB;
        i : integer;
        lastprix,mm : double;
begin
  // On augmente les lignes
for i := 0 to  TobPiece.detail.count-1 do
   BEGIN
     TobDet :=  TobPiece.detail[i];
     if (TobDet.getValue('GL_TYPELIGNE')='ART') and (TobDet.getvalue('GL_TYPEARTICLE')<>'POU') then
     Begin
       mm:=  TobDet.GetValue('GL_PUHTDEV');
       mm := CalcPourAugm(pour,mm,Annul);
       mm:=ArrondirPrix(ZArrondi,mm);
       mm:=Arrondi(mm,Dev.Decimale);


       lastprix := TobDet.GetValue('GL_PUHTDEV');
       TobDet.PutValue('GL_PUHTDEV', mm);
       if (Pos(TOBDet.GetValue('GL_TYPEARTICLE'),'OUV;ARP')>0) and (TOBDet.GetValue('GL_INDICENOMEN')<>0) then
       begin
       	TraitePrixOuvrage(TOBPiece,TOBDet,TOBBases,TOBBasesL,TOBOuvrage,(TOBDet.GetValue('GL_FACTUREHT')='X'),mm,lastprix,DEV);
       end;
		End;
   END;
//ProrateEchesGC( TOBPiece, TOBEches,Nil,DEV);
CalculPort;
Result := True;
END;

procedure T_ModifPieceAff.CalculPort ;
var Base, Total, Pourc, Montmini : Double ;
    TypePort : String ;
    TobPL : TOB;
    i :integer;
    SaisieContre: Boolean;
    X : double;
BEGIN
if  toBPorcs=NIL then exit;
for i := 0 to  TobPorcs.detail.count-1 do
   BEGIN
   TobPL :=  TobPorcs.detail[i];
   TypePort:=TobPl.GetValue('GPT_TYPEPORT') ;
   if ( TypePort<> 'MT') then  begin
      Pourc:=TOBPL.GetValue('GPT_POURCENT') ;
      Base:= TobPiece.Somme('GL_TOTALHTDEV',['GL_TYPELIGNE'],['ART'],False);
      Base:= Base -  TobPiece.GetValue('GP_TOTALESCDEV');
      TOBPL.PutValue('GPT_BASEHTDEV', BAse) ;
      Total:=Arrondi(Base*Pourc/100.0, DEV.Decimale);
       if TypePort= 'HT' then  begin
          TOBPL.PutValue('GPT_TOTALHTDEV', Total);
        END else if TypePort='MI' then
        BEGIN
          Montmini:=TOBPL.GetValue('GPT_MONTANTMINI');
          if Total < Montmini then Total:=Montmini ;
          if TOBPL.GetValue ('GPT_FRANCO')='X' then
              if Base >= TOBPL.GetValue ('GPT_MINIMUM') then Total:=0;
          TOBPL.PutValue('GPT_TOTALHTDEV', Total) ;
        END ;
    if  TobPiece.GetValue('GP_SAISIECONTRE')= 'X' then SaisieContre:=True
                     else SaisieContre:=False;
    If (DEV.Code <> V_PGI.DevisePivot) Then begin
        X:=TobPl.GetValue('GPT_TOTALHTDEV') ;
          TobPl.PutValue('GPT_TOTALHT',DeviseToEuro (X,DEV.Taux, DEV.Quotite)) ;
        X:=TobPl.GetValue('GPT_BASEHTDEV') ;
          TobPl.PutValue('GPT_BASEHT',DeviseToEuro (X,DEV.Taux, DEV.Quotite)) ;
      end
      else begin
        if Saisiecontre then
           begin
           TOBPL.PutValue('GPT_TOTALHT',Arrondi(DeviseToEuro(Total,DEV.Taux,DEV.Quotite),DEV.Decimale)) ;
           TOBPL.PutValue('GPT_BASEHT',Arrondi(DeviseToEuro(Base,DEV.Taux,DEV.Quotite),DEV.Decimale)) ;
           end
         else begin
          TOBPL.PutValue('GPT_TOTALHT',Arrondi(DeviseToFranc(Total,DEV.Taux,DEV.Quotite),DEV.Decimale)) ;
          TOBPL.PutValue('GPT_BASEHT',Arrondi(DeviseToFranc(Base,DEV.Taux,DEV.Quotite),DEV.Decimale)) ;
          end;
       end;
     end;
   end;
END;


Function T_ModifPieceAff.PassageEuro(AcbTenueEuro:boolean) : Boolean;
var     TobPrc, TobDet, TobEch : TOB;
        i : integer;
begin

  // Passage en Euro de la piece
ConvertEuro( TobPiece, 'PIECE', AcbTenueEuro );

  // Passage en Euro des lignes
for i := 0 to  TobPiece.detail.count-1 do
   BEGIN
   TobDet :=  TobPiece.detail[i];
   ConvertEuro( TobDet, 'LIGNE', AcbTenueEuro );
   END;

  // Passage en Euro des Pieds Porcs
for i := 0 to  TobPorcs.detail.count-1 do
   BEGIN
   TobPrc :=  TobPorcs.detail[i];
   ConvertEuro( TobPrc, 'PIEDPORT', AcbTenueEuro );
   END;

  // Passage en Euro des Pieds echeance
for i := 0 to  TobEches.detail.count-1 do
   BEGIN
   TobEch :=  TobEches.detail[i];
   ConvertEuro( TobEch, 'PIEDECHE', AcbTenueEuro );
   END;

Result := True;
END;


//******************************************************************************
// Fonction spécifique de recalcul des contrats pour ne pas générer d'erreurs si
// les enreg n'existe pas ( outils de recalcul des contrats ...)
// Attention : Non utilisé dans la GA
//******************************************************************************
{$IFDEF RECALCULAFF}
Function  DetruitAncienSansVerif ( TOBPiece_O,TOBBases_O,TOBEches_O,TOBN_O,TOBLOT_O,TOBAcomptes_O,TOBPorcs_O : TOB ) : boolean ;
Var CleDoc : R_CleDoc ;
    Nb     : integer ;
    RefA   : String ;
    OldD   : TDateTime ;
BEGIN
Result:=False ;
CleDoc:=TOB2CleDoc(TOBPiece_O) ; OldD:=TOBPiece_O.GetValue('GP_DATEMODIF') ;
Nb:=ExecuteSQL('DELETE FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False)+' AND GP_DATEMODIF="'+USTime(OldD)+'"') ;
if Nb<=0 then BEGIN V_PGI.IoError:=oeSaisie ; Exit ; END ;
if TOBPiece_O.Detail.Count>0 then
   BEGIN
   Nb:=ExecuteSQL('DELETE FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False)) ;
   if Nb<=0 then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
   END ;
if TOBBases_O.Detail.Count>0 then
   BEGIN
   Nb:=ExecuteSQL('DELETE FROM PIEDBASE WHERE '+WherePiece(CleDoc,ttdPiedBase,False)) ;
   // if Nb<=0 then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
   END ;
if TOBEches_O.Detail.Count>0 then
   BEGIN
   Nb:=ExecuteSQL('DELETE FROM PIEDECHE WHERE '+WherePiece(CleDoc,ttdEche,False)) ;
   // if Nb<=0 then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
   END ;
if TOBN_O.Detail.Count>0 then
   BEGIN
   Result:=TOBN_O.DeleteDB ;
   if Not Result then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
   END ;
if TOBLOT_O<>Nil then if TOBLOT_O.Detail.Count>0 then
   BEGIN
   Result:=TOBLOT_O.DeleteDB ;
   if Not Result then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
   END ;
if TOBAcomptes_O<>Nil then if TOBAcomptes_O.Detail.Count>0 then
   BEGIN
   Result:=TOBAcomptes_O.DeleteDB ;
   if Not Result then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
   END ;
if TOBPorcs_O<>Nil then if TOBPorcs_O.Detail.Count>0 then
   BEGIN
   Result:=TOBPorcs_O.DeleteDB ;
   if Not Result then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
   END ;
// RefA:=EncodeRefCPGescom(TOBPiece_O) ; Nb:=ExecuteSQL('DELETE FROM VENTANA WHERE YVA_TABLEANA="GL" AND YVA_IDENTIFIANT="'+RefA+'"') ;
Result:=True ;
END ;
{$ENDIF}

//******************************************************************************
// FIN   Fonction spécifique de recalcul des contrats
// Attention : Non utilisé dans la GA
//******************************************************************************




//******************  Validation ***********************************************
Procedure T_ModifPieceAff.ValideModifPiece ;
var Ind : integer;
BEGIN   // fct ValideLaPiece de Facture.pas
InitToutModif;
ValideLaCotation(TOBPiece,TOBBases,TOBEches) ;
ValideLaPeriode(TOBPiece) ;

{Calculs et finitions}
MajLesComms(TOBPiece) ;
//modif BTP
  ZeroFacture (TOBpiece);
  for Ind := 0 to TOBPiece.detail.count -1 do ZeroLigneMontant (TOBPiece.detail[Ind]);
  PutValueDetail (TOBpiece,'GP_RECALCULER','X');
  TOBBases.ClearDetail;
  TOBBasesL.ClearDetail;
//---
CalculFacture(nil,TOBPiece,TOBPieceTrait,TOBSSTRAIT, TOBOuvrage,TOBOuvragesP,TOBBases,TobbasesL,TOBTiers,TOBArticles,TOBPorcs,Nil,Nil,nil,DEV) ;
CalculeSousTotauxPiece(TOBPiece) ;

{Détruit la pièce pour la réécrire}
{$IFDEF RECALCULAFF}
// PA - spécif recalcul des affaires/contrat pour ne pas planter si les bases, eche n'existent pas
DetruitAncienSansVerif (TOBPiece,TOBBases,TOBEches,TOBNomenclature,Nil,TOBAcomptes,TOBPorcs);
{$ELSE}
DetruitAncien (TOBPiece,TOBBases,TOBEches,TOBNomenclature,Nil,TOBAcomptes,TOBPorcs,Nil,TOBOuvrage,Nil,Nil,NIl, nil);
{$ENDIF}

{Enregistrement physique}
if V_PGI.IoError=oeOk then ValideLesLignes(TOBPiece,TOBArticles,TobCatalogu,TOBNomenclature,TOBouvrage,Nil,NIl,False,True) ;
if V_PGI.IoError = oeOk then ValideLesBases(TOBPiece,TobBases,TOBBasesL);
//if V_PGI.IoError=oeOk then ValideLesAdresses(TOBPiece,TOBPiece,TOBAdresses) ;
{$IFNDEF RECALCULAFF}
if V_PGI.IoError=oeOk then ValideLesArticles(TOBPiece,TOBArticles) ;
if V_PGI.IoError=oeOk then ValideLesCatalogues(TOBPiece,TOBCatalogu) ;
// if V_PGI.IoError=oeOk then G_ValideAnals ;
// if V_PGI.IoError=oeOk then G_GenereCompta ;
if V_PGI.IoError=oeOk then ValideLesAcomptes(TOBPiece,TOBAcomptes) ;
if V_PGI.IoError=oeOk then ValideLesPorcs(TOBPiece,TOBPorcs) ;
{$ENDIF}
// 25/11/02 if V_PGI.IoError=oeOk then TOBpiece.InsertDBByNivelTable(False) ;
if V_PGI.IoError=oeOk then TOBpiece.InsertDBByNivel(False) ;
(* mcd 25/11/02 ne plus utiliser ...if V_PGI.IoError=oeOk then TOBBases.InsertDBTable(Nil) ;
if V_PGI.IoError=oeOk then TOBEches.InsertDBTable(Nil) ;
if V_PGI.IoError=oeOk then TOBAnaP.InsertDBTable(Nil) ;
if V_PGI.IoError=oeOk then TOBAnaS.InsertDBTable(Nil) ;
if V_PGI.IoError=oeOk then TOBBases.InsertDBTable(Nil) ;*)
if V_PGI.IoError=oeOk then TOBEches.InsertDB(Nil) ;
if V_PGI.IoError=oeOk then TOBAnaP.InsertDB(Nil) ;
if V_PGI.IoError=oeOk then TOBAnaS.InsertDB(Nil) ;
if V_PGI.IoError=oeOk then TOBBases.InsertDB(Nil) ;
if V_PGI.IoError=oeOk then ValideLesNomen(TOBNomenclature) ;
if V_PGI.IoError=oeOk then ValideLesOuv(TOBOuvrage, TOBPiece);

END ;

Procedure T_ModifPieceAff.MajLesComms ( TOBPiece : TOB ) ;
Var i : integer ;
    EtatVC : String ;
    TOBL : TOB ;
BEGIN
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ;
    EtatVC:=TOBL.GetValue('GL_VALIDECOM') ;
    if EtatVC='VAL' then TOBL.PutValue('GL_VALIDECOM','AFF') else
     if EtatVC='AFF' then CommVersLigne(TOBPiece,TOBArticles,TOBComms,i+1,False) ;
    END ;
END ;


(*******************************************************************************
Fonctions de conversion en Euro
*******************************************************************************)
Procedure BourreLeDEV ( TobF : TOB ; NomChamp : String ; bParLeCon:boolean) ;
BEGIN
if Not BparLeCon then TobF.PutValue(NomChamp+'DEV', TobF.GetValue(NomChamp));
END ;

Procedure BourreLeChamp ( TobF : TOB ; NomChampOld, NomChampNew : String ) ;
BEGIN
TobF.PutValue(NomChampOld, TobF.GetValue(NomChampNew));
END ;

Procedure SwapLeFlag ( TobF : TOB ; NomChamp : String ) ;
Var St : String ;
BEGIN
St:=TobF.GetValue(NomChamp);
if St='-' then St:='X' else St:='-' ;
TobF.PutValue(NomChamp, St);
END ;

Procedure ConvertEuro ( TobF : TOB ; NomT : String; bTenueEuro:boolean );
Var i    : integer ;
bParLeCon:boolean;
BEGIN
bParLeCon := not bTenueEuro;
if NomT='' then Exit ;
if NomT='PIECE' then
   BEGIN
   BourreLeDEV(TobF,'GP_TOTALESC',bParLeCon)     ; BourreLeDEV(TobF,'GP_TOTESCTTC',bParLeCon) ;
   BourreLeDEV(TobF,'GP_TOTALREMISE',bParLeCon)  ; BourreLeDEV(TobF,'GP_TOTREMISETTC',bParLeCon) ;
   BourreLeDEV(TobF,'GP_TOTALHT',bParLeCon)      ; BourreLeDEV(TobF,'GP_TOTALTTC',bParLeCon) ;
   BourreLeDEV(TobF,'GP_TOTALBASEREM',bParLeCon) ; BourreLeDEV(TobF,'GP_TOTALBASEESC',bParLeCon) ;
   BourreLeDEV(TobF,'GP_ACOMPTE',bParLeCon) ;
   SwapLeFlag(TobF,'GP_SAISIECONTRE') ;
   END else
if NomT='PIEDBASE' then
   BEGIN
    if Not bParLeCon then
     begin
     BourreLeChamp(TobF,'GPB_BASEDEV','GPB_BASETAXE');
     BourreLeChamp(TobF,'GPB_VALEURDEV','GPB_VALEURTAXE');
     end;

   SwapLeFlag(TobF,'GPB_SAISIECONTRE') ;
   END else
if NomT='PIEDECHE' then
   BEGIN
   if Not bParLeCon then BourreLeChamp(TobF,'GPE_MONTANTDEV','GPE_MONTANTECHE');
   SwapLeFlag(TobF,'GPE_SAISIECONTRE') ;
   END else
if NomT='PIEDPORT' then
   BEGIN
   BourreLeDEV(TobF,'GPT_BASEHT',bParLeCon)  ; BourreLeDEV(TobF,'GPT_BASETTC',bParLeCon) ;
   BourreLeDEV(TobF,'GPT_TOTALHT',bParLeCon) ; BourreLeDEV(TobF,'GPT_TOTALTTC',bParLeCon) ;
   for i:=1 to 5 do
        if Not bParLeCon then BourreLeChamp(TobF,'GPT_TOTALTAXEDEV'+IntToStr(i),'GPT_TOTALTAXE'+IntToStr(i));
   END else
if NomT='LIGNE' then
   BEGIN
   BourreLeDEV(TobF,'GL_TOTALHT',bParLeCon)     ; BourreLeDEV(TobF,'GL_TOTALTTC',bParLeCon) ;
   BourreLeDEV(TobF,'GL_MONTANTHT',bParLeCon)   ; BourreLeDEV(TobF,'GL_MONTANTTTC',bParLeCon) ;
   BourreLeDEV(TobF,'GL_TOTREMPIED',bParLeCon)  ; BourreLeDEV(TobF,'GL_TOTREMLIGNE',bParLeCon) ;
   BourreLeDEV(TobF,'GL_PUHT',bParLeCon)        ; BourreLeDEV(TobF,'GL_PUTTC',bParLeCon) ;
   BourreLeDEV(TobF,'GL_PUHTNET',bParLeCon)     ; BourreLeDEV(TobF,'GL_PUTTCNET',bParLeCon) ;
   BourreLeDEV(TobF,'GL_TOTESCLIGNE',bParLeCon) ; BourreLeDEV(TobF,'GL_PUHTORIGINE',bParLeCon) ;
   for i:=1 to 5 do
        if Not bParLeCon  then BourreLeChamp(TobF,'GL_TOTALTAXEDEV'+IntToStr(i),'GL_TOTALTAXE'+IntToStr(i)) ;
   SwapLeFlag(TobF,'GL_SAISIECONTRE') ;
   END
else
if NomT='AFFAIRE' then
   BEGIN
   BourreLeDEV(TobF,'AFF_REPORT',bParLeCon)  ;
   //BourreLeDEV(TobF,'AFF_TOTALHT')  ;
   //BourreLeDEV(TobF,'AFF_TOTALTTC')  ;
   //BourreLeDEV(TobF,'AFF_TOTALTAXE')  ;
   if (TobF.Getvalue('AFF_CALCTOTHTGLO')='-') then
        BourreLeDEV(TobF,'AFF_TOTALHTGLO',bParLeCon)  ;
   BourreLeDEV(TobF,'AFF_MONTANTECHE',bParLeCon)  ;
   SwapLeFlag(TobF,'AFF_SAISIECONTRE') ;
   END else
if NomT='FACTAFF' then
   BEGIN
   if (TobF.GetValue('AFA_TYPECHE')='NOR') then
        BourreLeDEV(TobF,'AFA_MONTANTECHE',bParLeCon)
   else
        begin
        BourreLeDEV(TobF,'AFA_BONIMALI',bParLeCon)  ;
        BourreLeDEV(TobF,'AFA_AFACTURER',bParLeCon)  ;
        BourreLeDEV(TobF,'AFA_BM1FO',bParLeCon)  ;
        BourreLeDEV(TobF,'AFA_BM1FR',bParLeCon)  ;
        BourreLeDEV(TobF,'AFA_BM2PR',bParLeCon)  ;
        BourreLeDEV(TobF,'AFA_BM2FO',bParLeCon)  ;
        BourreLeDEV(TobF,'AFA_BM2FR',bParLeCon)  ;
        BourreLeDEV(TobF,'AFA_AFACTOT',bParLeCon)  ;
        BourreLeDEV(TobF,'AFA_AFACTFR',bParLeCon)  ;
        BourreLeDEV(TobF,'AFA_AFACTFO',bParLeCon)  ;
        BourreLeDEV(TobF,'AFA_ACPTEPR',bParLeCon)  ;
        BourreLeDEV(TobF,'AFA_ACPTEFO',bParLeCon)  ;
        BourreLeDEV(TobF,'AFA_ACPTEFR',bParLeCon)  ;
        end;
   END;

END ;

Function CalcPourAugm(pour,mm : double;Annul:Boolean) : double;
BEGIN
	if (pour > 0) then
  	result  := mm * (1+pour)
  else
  	begin    // si annul vrai, on veut effacer une précédente augmentation, avec le même %. sinon, calcul réel du % en -
    	If Annul then result := mm / (1+(pour*(-1)))
           else  result  := mm * (1+pour)
    end;
END;

end.
