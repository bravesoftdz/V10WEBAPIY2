unit FactActivite;

interface

uses UTOB,SysUtils,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HCtrls,Hent1,Ent1,classes, ActiviteUtil, FactUtil,EntGC,UtofPrixRevient,AffaireUtil
     ,TiersUtil,ParamSoc,factcomm,utilressource;


// fonctions de gestion de l'activité
procedure EtudieActivite(NaturePiece : String;Action : TActionFiche;duplicpiece : boolean; Var GereActivite,DelActivite : Boolean) ;
procedure ValideActivite (TobPiece, Tobpiece_O,TobArticles : TOB; GereActivite,Reliquat: Boolean ; Var DelActivite : Boolean; MajDel : Boolean = True);
procedure PieceToActivite(TOBPiece, TobArticles : TOB; GereActivite,Reliquat : Boolean);
function  LigneToAct (TobLigne, TobAct, TobArticles : TOB; Reliquat : Boolean; ListeDesCles : TStringList) : Boolean;
function  EtudieLienligneToAct (TobLigne : TOB): Boolean;
Procedure ImpactActiviteLigne(TOBPiece : TOB; GereActivite : Boolean);
// Valorisation des lignes d'activités générées
Procedure ValoriseLigneAct ( TobLigne, TobDetAct, TobArt : TOB);
Procedure CalculPVLigneAct (TobLigne, TobDetAct, TobArt : TOB);
Procedure CalculPRLigneAct (TobLigne, TobDetAct, TobArt : TOB);
// gestion des clés d'activités
//Function  GetNumLigneInCleAct(TobLigneAct, TobCleAct : TOB) : integer;
//Procedure InsertCleInCleAct(TobLigneAct, TobCleAct : TOB ; NumLigne : Integer);

implementation


procedure EtudieActivite(NaturePiece : String;Action : TActionFiche;duplicpiece : boolean; Var GereActivite,DelActivite : Boolean) ;
BEGIN
GereActivite :=True;
if not(ctxAffaire in V_PGI.PGIContexte) And not(ctxGCAff in V_PGI.PGIContexte) then GereActivite := False;

if GereActivite then
   GereActivite:=((Action=taModif) or (Action=taCreat)) and
                 (GetInfoParPiece(NaturePiece,'GPP_ACHATACTIVITE')='X');
if duplicpiece then
// en dupplication de pièce , il ne faut pas supprimer l'activite de la piece d'origine
  DelActivite := false
else
  DelActivite := True;
END;

procedure ValideActivite (TobPiece, TobPiece_O, TobArticles : TOB; GereActivite,Reliquat : Boolean; Var DelActivite : Boolean; majDel : Boolean = True);
Var  NumPiece : integer;
    stSQL : string;
BEGIN
if not(ctxAffaire in V_PGI.PGIContexte) and not(ctxGCAFF in V_PGI.PGIContexte) then Exit;
        // si option facture éclatée par assistant, on appel fct adéquate
if (GetParamSoc('SO_AFFACTPARRES') <> 'SAN')
   then  begin
   if (TobPiece <> NIL)  and
   ( (TOBPIECE.GetValue('GP_NATUREPIECEG') = 'FAC') or
     (TOBPIECE.GetValue('GP_NATUREPIECEG') = 'FRE') or
     (TOBPIECE.GetValue('GP_NATUREPIECEG') = 'AVC') )
      then MajFactEclat (TobPiece);
   end;


if (V_PGI.IoError<>oeOk)  then Exit;
if Not(GereActivite) then Exit;
// Suppression de l'activité sur la pièce précédente
if (TOBPiece_O <> Nil) And (DelActivite) then
   BEGIN
   if (TobPiece_O.Detail.Count <> 0)  then
      BEGIN
      NumPiece := TobPiece_O.GetValue('GP_NUMERO');
      stSQL := 'DELETE from ACTIVITE where ACT_NATUREPIECEG="' + TobPiece_O.GetValue('GP_NATUREPIECEG')+
               '" And ACT_SOUCHE="'+ TobPiece_O.GetValue('GP_SOUCHE') +
               '" And ACT_NUMERO=' + IntToStr(NumPiece) +
               '  And ACT_INDICEG='+ IntToStr(TobPiece_O.GetValue('GP_INDICEG'));
      if (NumPiece <> 0) then BEGIN ExecuteSQL (stSQL); if MajDel then DelActivite:=False; END;
      END;
   END;
// Création de l'activité sur la nouvelle pièce
if TobPiece <> Nil then
   PieceToActivite (Tobpiece,TobArticles,GereActivite,Reliquat);
END;


{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 23/11/2000
Modifié le ... : 23/11/2000
Description .. : passage d'une pièce d'achats en ligne d'activité
Mots clefs ... : ACTIVITE;ACHAT
*****************************************************************}
procedure PieceToActivite(TOBPiece, TobArticles : TOB; GereActivite,Reliquat : Boolean);
Var TobDet, TobAct : TOB;
    ListeDesCles : TStringList;
    i : integer;

BEGIN
if not(GereActivite) then Exit;
if TOBPiece.Detail.Count = 0 then exit;

TobAct := TOB.Create ('Reliquats activite',Nil,-1);
//TobCleAct := TOB.Create ('liste cles act' , Nil,-1);
ListeDesCles := TStringList.Create;
try

  for i := 0 to TOBPiece.Detail.Count-1 do
    BEGIN
      TOBDet := TOBPiece.Detail[i];
      if TobDet.GetValue('GL_ARTICLE') = '' then continue;
      LigneToAct (TobDet, TobAct, TobArticles, Reliquat, ListeDesCles);
    END;

  if TobAct.Detail.count > 0 then TobAct.InsertDB (Nil);

finally
  TobAct.Free;
  //TobCleAct.Free;
  ListeDesCles.Free;
end;

END;


{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 23/11/2000
Modifié le ... : 23/11/2000
Description .. : passage d'une ligne de pièce en ligne d'activité
Mots clefs ... : ACTIVITE;ACHAT
*****************************************************************}
function LigneToAct (TobLigne, TobAct, TobArticles : TOB; Reliquat : Boolean; ListeDesCles : TStringList): boolean;
Var NumLigne : integer;
    Q : TQuery;
    TobArt, TobDetAct : TOB;
BEGIN
Result := false;
if (TobLigne = Nil) or (TobAct = Nil) then Exit;
if Not(EtudieLienligneToAct (TobLigne)) then Exit;

TobDetAct := TOB.Create ('ACTIVITE', TobAct,-1);

TOBArt:=TOBArticles.FindFirst(['GA_ARTICLE'],[TobLigne.GetValue('GL_ARTICLE')],False) ;
// clés
TobDetAct.PutValue('ACT_TYPEACTIVITE', 'REA');
TobDetAct.PutValue('ACT_DATECREATION',Date);
TobDetAct.PutValue('ACT_FOLIO', 01);
TobDetAct.PutValue('ACT_ACTORIGINE', 'ACH');
TobDetAct.PutValue('ACT_DATEACTIVITE', TobLigne.GetValue('GL_DATEPIECE'));
TobDetAct.PutValue('ACT_PERIODE', GetPeriode(TobLigne.GetValue('GL_DATEPIECE')));
TobDetAct.PutValue('ACT_SEMAINE', NumSemaine(TobLigne.GetValue('GL_DATEPIECE')));
if (GetInfoParPiece(TobLigne.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT')='ACH') then
   TobDetAct.PutValue('ACT_TIERS',GetChampsAffaire(TobLigne.GetValue('GL_AFFAIRE'),'AFF_TIERS'))  //tiers ligne = fournisseur non ok...
else
   TobDetAct.PutValue('ACT_TIERS', TobLigne.GetValue('GL_TIERS')); // tiers ok = Client
TobDetAct.PutValue('ACT_AFFAIRE', TobLigne.GetValue('GL_AFFAIRE'));
TobDetAct.PutValue('ACT_AFFAIRE0', Copy(TobLigne.GetValue('GL_AFFAIRE'),1,1));
TobDetAct.PutValue('ACT_AFFAIRE1', TobLigne.GetValue('GL_AFFAIRE1'));
TobDetAct.PutValue('ACT_AFFAIRE2', TobLigne.GetValue('GL_AFFAIRE2'));
TobDetAct.PutValue('ACT_AFFAIRE3', TobLigne.GetValue('GL_AFFAIRE3'));
TobDetAct.PutValue('ACT_AVENANT', TobLigne.GetValue('GL_AVENANT'));

TobDetAct.PutValue('ACT_TYPEARTICLE', TobLigne.GetValue('GL_TYPEARTICLE'));
TobDetAct.PutValue('ACT_ARTICLE', TobLigne.GetValue('GL_ARTICLE'));
TobDetAct.PutValue('ACT_CODEARTICLE', TobLigne.GetValue('GL_CODEARTICLE'));
TobDetAct.PutValue('ACT_LIBELLE', TobLigne.GetValue('GL_LIBELLE'));
//mcd 25/09/03 ?? pas bonne zone ?TobDetAct.PutValue('ACT_UNITE', TobLigne.GetValue('GL_QUALIFQTEACH')); // ou stock ???
TobDetAct.PutValue('ACT_UNITE', TobLigne.GetValue('GL_QUALIFQTEVTE')); // ou stock ???
TobDetAct.PutValue('ACT_QTE', TobLigne.GetValue('GL_QTEFACT'));
//mcd 25/09/03 TobDetAct.PutValue('ACT_UNITEFAC', TobLigne.GetValue('GL_QUALIFQTEACH'));
TobDetAct.PutValue('ACT_UNITEFAC', TobLigne.GetValue('GL_QUALIFQTEVTE'));
if (TOBDetAct.GetValue ('ACT_TYPEARTICLE') = 'PRE') then
  TOBDetAct.PutValue ('ACT_QTEUNITEREF', ConversionUnite (TOBDetAct.GetValue ('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TOBDetAct.GetValue ('ACT_QTE'))); //mcd 16/09/03
TobDetAct.PutValue('ACT_QTEFAC', TobLigne.GetValue('GL_QTEFACT'));
if (TobDetAct.GetValue ('ACT_TYPEARTICLE') = 'PRE') then
  TobDetAct.PutValue ('ACT_QTEUNITEREF', ConversionUnite (TobDetAct.GetValue ('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TobDetAct.GetValue('ACT_QTE')))
else
  TobDetAct.PutValue ('ACT_QTEUNITEREF', '0');

TobDetAct.PutValue('ACT_DEVISE',V_PGI.DevisePivot);      // repris en monnaie pivot pour l'instant  TobLigne.GetValue('GL_DEVISE')

// Valorisation de la ligne d'activité
ValoriseLigneAct ( TobLigne, TobDetAct, TobArt);

// Repris de l'article
if TobArt <> Nil then TobDetAct.PutValue('ACT_ACTIVITEREPRIS', TobArt.GetValue('GA_ACTIVITEREPRISE'));
// mcd 26/05/03 ??? test à l'envers ??? if TobDetAct.GetValue('ACT_ACTIVITEREPRIS') <> '' then TobDetAct.PutValue('ACT_ACTIVITEREPRIS', 'F');
if TobDetAct.GetValue('ACT_ACTIVITEREPRIS') = '' then TobDetAct.PutValue('ACT_ACTIVITEREPRIS', 'F');

// traitement ressource
TobDetAct.PutValue('ACT_RESSOURCE', TobLigne.GetValue('GL_RESSOURCE'));
if TobLigne.GetValue('GL_RESSOURCE') <> '' then
   BEGIN
   Q:= OPENSQL ('SELECT ARS_TYPERESSOURCE,ARS_FONCTION1 FROM RESSOURCE WHERE ARS_RESSOURCE="'+
                TobLigne.GetValue('GL_RESSOURCE')+'"',True,-1, '', True);
   if Not(Q.EOF) then
      BEGIN
      TobDetAct.PutValue('ACT_TYPERESSOURCE', Q.Fields[0].AsString);
      TobDetAct.PutValue('ACT_FONCTIONRES', Q.Fields[1].AsString);
      END;
   Ferme(Q);
   END
else begin
//  if (TobDetAct.GetValue('ACT_TYPEARTICLE')='PRE') or (TobDetAct.GetValue('ACT_TYPEARTICLE')='FRA')
//     then begin // cas ou on traite aussi frais et fourniture, il faut renseigner la ressource
     Q:= OPENSQL ('SELECT ARS_RESSOURCE,ARS_TYPERESSOURCE,ARS_FONCTION1 FROM RESSOURCE WHERE ARS_AUXILIAIRE="'+
                  TiersAuxiliaire(TobLigne.GetValue('GL_TIERS'),False)+'"',True,-1, '', True);
     if Not(Q.EOF) then
        BEGIN
        TobDetAct.PutValue('ACT_TYPERESSOURCE', Q.FindField('ARS_TYPERESSOURCE').asString);
        TobDetAct.PutValue('ACT_RESSOURCE', Q.FindField('ARS_RESSOURCE').AsString);
        TobDetAct.PutValue('ACT_FONCTIONRES', Q.FindField('ARS_FONCTION1').AsString);
        END
        else begin
        Ferme(Q);
        TobDetAct.PutValue('ACT_RESSOURCE', GetParamSoc('SO_AFRESS_DEF_FOUR'));
        Q:= OPENSQL ('SELECT ARS_TYPERESSOURCE,ARS_FONCTION1 FROM RESSOURCE WHERE ARS_RESSOURCE="'+
                  GetParamSoc('SO_AFRESS_DEF_FOUR')+'"',True,-1, '', True);
        if Not(Q.EOF) then
            BEGIN
            TobDetAct.PutValue('ACT_TYPERESSOURCE', Q.FindField('ARS_TYPERESSOURCE').AsString);
            TobDetAct.PutValue('ACT_FONCTIONRES', Q.FindField('ARS_FONCTION1').AsString);
            end;
        end;
     Ferme(Q);
     end;
//     end;

// Num de pièce associée
TobDetAct.PutValue('ACT_NATUREPIECEG', TobLigne.GetValue('GL_NATUREPIECEG'));
TobDetAct.PutValue('ACT_SOUCHE', TobLigne.GetValue('GL_SOUCHE'));
TobDetAct.PutValue('ACT_NUMERO', TobLigne.GetValue('GL_NUMERO'));
TobDetAct.PutValue('ACT_INDICEG', TobLigne.GetValue('GL_INDICEG'));
TobDetAct.PutValue('ACT_FOURNISSEUR', TobLigne.GetValue('GL_TIERS'));

// gm 14/3/2002 Ajout alim numpiece
TobDetAct.PutValue('ACT_NUMPIECEACH',EncodeRefPiece(Tobligne));

TobDetAct.PutValue('ACT_DESCRIPTIF', TobLigne.GetValue('GL_BLOCNOTE'));
// gm 03/05/2002 Ajout notion visa
// Gestion automatique du visa sur l'activité
  if (GetParamSoc('SO_AFVISAACTIVITE')=false) then
  begin
    TobDetAct.PutValue('ACT_ETATVISA', 'VIS');
    TobDetAct.PutValue('ACT_VISEUR', V_PGI.User);
    TobDetAct.PutValue('ACT_DATEVISA', NowH);
  end
  else
  begin
    TobDetAct.PutValue('ACT_ETATVISA', 'ATT');
    TobDetAct.PutValue('ACT_DATEVISA', idate1900);
  end;

  // Gestion automatique du visa de facturation sur l'activité
  if (GetParamSoc('SO_AFAPPPOINT')=false) then
  begin
    TobDetAct.PutValue('ACT_ETATVISAFAC', 'VIS');
    TobDetAct.PutValue('ACT_VISEURFAC', V_PGI.User);
    TobDetAct.PutValue('ACT_DATEVISAFAC', NowH);
  end
  else
  begin
   TobDetAct.PutValue('ACT_ETATVISAFAC', 'ATT');
   TobDetAct.PutValue('ACT_DATEVISAFAC', idate1900);
  end;



TobDetAct.PutValue('ACT_ETATVISAFAC', 'VIS');
TobDetAct.PutValue('ACT_VISEURFAC', V_PGI.User);
TobDetAct.PutValue('ACT_DATEVISAFAC', NowH);

///////////// PL le 14/04/03 : modifs changement de clé dans la table ACTIVITE
// Recherche du num de ligne de la clé 1
(*NumLigne := GetNumLigneInCleAct(TobDetAct, TobCleAct);
if NumLigne = 0 then
   BEGIN
   NumLigne := MaxNumLigneActivite(TobDetAct); // Attention une requête est faite à chaque fois ...
   InsertCleInCleAct  (TobDetAct,TobCleAct,NumLigne+1);
   END;
TobDetAct.PutValue('ACT_NUMLIGNE', NumLigne+1);*)
NumLigne := ProchainIndiceAffaires (TobDetAct.GetValue ('ACT_TYPEACTIVITE'), TobDetAct.GetValue ('ACT_AFFAIRE'), ListeDesCles);
TobDetAct.PutValue ('ACT_NUMLIGNEUNIQUE', NumLigne);
TobDetAct.PutValue ('ACT_NUMLIGNE', NumLigne); // La valeur du numéro d'ordre a peu d'importance
//////////////// Fin PL le 14/04/03


// Attention si une pers. intègre une ligne sur la même affaire / date ???
Result := True;
END;


function  EtudieLienligneToAct (TobLigne : TOB): Boolean;
BEGIN
Result := False;
if TobLigne.GetValue('GL_ARTICLE') = '' then Exit;
if TobLigne.GetValue('GL_AFFAIRE') = '' then Exit;

Result := True;
END;

//******************************************************************************
//******************* Gestion des clés de la table Activité ********************
//******************************************************************************
(*Function GetNumLigneInCleAct (TobLigneAct, TobCleAct : TOB) : integer;
  Var
  Cle : string;
  TobDet : TOB;
BEGIN
  Result := 0;
  if (TobCleAct = Nil) or (TobLigneAct = Nil) then Exit;
  // PL le 16/04/03 : modif clé ACTIVITE
  //Cle:=TobLigneAct.GetValue('ACT_TYPEACTIVITE')+';'+TobLigneAct.GetValue('ACT_AFFAIRE')+';'
  //                +TobLigneAct.GetValue('ACT_RESSOURCE')+';'+DateToStr(TobLigneAct.GetValue('ACT_DATEACTIVITE'))+';'
  //                +TobLigneAct.GetValue('ACT_TYPEARTICLE');
  Cle := TobLigneAct.GetValue ('ACT_TYPEACTIVITE') + ';' + TobLigneAct.GetValue ('ACT_AFFAIRE') + ';';
  TobDet := TobCleact.FindFirst (['CLEACT'], [Cle], False);
  if TobDet <> Nil then
    BEGIN
      Result := TobDet.GetValue ('NUMLIGNE');
      TobDet.PutValue ('NUMLIGNE', TobDet.GetValue ('NUMLIGNE') + 1);
    END;
  if Result < 0 then Result := 0;
END;

Procedure InsertCleInCleAct (TobLigneAct, TobCleAct : TOB; NumLigne : Integer);
  Var
  Cle : String;
  TobDet : TOB;
BEGIN
  if TobCleAct = Nil then Exit;
  // PL le 16/04/03 : modif clé ACTIVITE
  //Cle := TobLigneAct.GetValue('ACT_TYPEACTIVITE')+';'+TobLigneAct.GetValue('ACT_AFFAIRE')+';'
  //                +TobLigneAct.GetValue('ACT_RESSOURCE')+';'+DateToStr(TobLigneAct.GetValue('ACT_DATEACTIVITE'))+';'
  //                +TobLigneAct.GetValue('ACT_TYPEARTICLE');
  Cle := TobLigneAct.GetValue ('ACT_TYPEACTIVITE') + ';' + TobLigneAct.GetValue ('ACT_AFFAIRE') + ';';
  // fin  PL le 16/04/03 : modif clé ACTIVITE
  TobDet := TobCleact.FindFirst (['CLEACT'], [Cle], False);
  if TOBDet <> Nil then Exit;
  TobDet := TOB.create ('Cle activite', TobCleAct, -1);
  TobDet.AddChampSup ('CLEACT', false);
  TobDet.PutValue ('CLEACT', Cle);
  TobDet.AddChampSup ('NUMLIGNE', false);
  TobDet.PutValue ('NUMLIGNE', NumLigne);
END;
*)

//******************************************************************************
//******************* Valorisation d'une ligne d'activité   ********************
//******************************************************************************
Procedure ValoriseLigneAct ( TobLigne, TobDetAct, TobArt : TOB);
Var PuPR,TotPR : Double;
BEGIN
// **** Calcul de la valorisation de la ligne d'activité ****
if (GetInfoParPiece(TobLigne.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT')='ACH') then
   begin
   // pièce d'achat
   PUPR  := TobLigne.GetValue('GL_PUHT');             // Prix d'achat de la ligne en monnaie pivot
   TotPR := TobLigne.GetValue('GL_TOTALHT');
      // PUPR = Prix d'achat de la ligne
   TobDetAct.PutValue('ACT_PUPR', PUPR);
   TobDetAct.PutValue('ACT_TOTPR', TOTPR);
      // PUPRCHARGE = Prix de revient de l'article suivant les règles de calcul des PR
   CalculPRLigneAct (TobLigne, TobDetAct, TobArt); // atention après calcul ACT_PUPR car repris
      // Idem PUPRCharge
   TobDetAct.PutValue('ACT_PUPRCHINDIRECT', TobDetAct.GetValue('ACT_PUPRCHARGE'));
   TobDetAct.PutValue('ACT_TOTPRCHINDI', TobDetAct.GetValue('ACT_TOTPRCHARGE'));
      // Prix de vente - recalcul par rapport au coef PR / PV de l'article...
   CalculPVLigneAct (TobLigne, TobDetAct, TobArt);
   end
else
   begin
   // pièce de vente
      // prix de revient
   TobDetAct.PutValue('ACT_PUPR', Tobligne.GetValue('GL_DPA'));
      // PUPRCHARGE = Prix de revient de l'article suivant les règles de calcul des PR
   CalculPRLigneAct (TobLigne, TobDetAct, TobArt); // atention après calcul ACT_PUPR car repris
      // Idem PUPRCharge
   TobDetAct.PutValue('ACT_PUPRCHINDIRECT', TobDetAct.GetValue('ACT_PUPRCHARGE'));
   TobDetAct.PutValue('ACT_TOTPRCHINDI', TobDetAct.GetValue('ACT_TOTPRCHARGE'));
      // prix de vente
   TobDetAct.PutValue('ACT_PUVENTE', TobLigne.GetValue('GL_PUHT'));
   TobDetAct.PutValue('ACT_PUVENTEDEV', TobLigne.GetValue('GL_PUHTDEV'));
   TobDetAct.PutValue('ACT_TOTVENTE', TobLigne.GetValue('GL_TOTALHT'));
   TobDetAct.PutValue('ACT_TOTVENTEDEV', TobLigne.GetValue('GL_TOTALHTDEV'));
   end;
END;

Procedure CalculPVLigneAct (TobLigne, TobDetAct, TobArt : TOB);
Var CodeBaseHT,CodeArrondi : String;
    CoefCalcHT,PrixRef,BaseTarifHT,TotTarifHT : Double;
BEGIN
if (TobArt = Nil) Then Exit;
PrixRef:=0;

CodeBaseHT:=TobArt.GetValue('GA_CALCPRIXHT');
CoefCalcHT:=TobArt.GetValue('GA_COEFCALCHT');
CodeArrondi:=TobArt.GetValue('GA_ARRONDIPRIX');

if (CoefCalcHT<>0) and (CodeBaseHT<>'') then
  begin
  if CodeBaseHT='DPA' then PrixRef:=TobLigne.GetValue('GL_DPA') else
  if CodeBaseHT='DPR' then PrixRef:=TobLigne.GetValue('GL_DPR') else
  if CodeBaseHT='PAA' then PrixRef:=TobLigne.GetValue('GL_PUHT')else
  if CodeBaseHT='PMA' then PrixRef:=TobArt.GetValue('GA_PMAP')else
  if CodeBaseHT='PMR' then PrixRef:=TobArt.GetValue('GA_PMRP');
  BaseTarifHT:=Arrondi(PrixRef*CoefCalcHT,V_PGI.OkDecV) ;
  end
else
   BaseTarifHT:=TobArt.GetValue('GA_PVHT');

  TotTarifHT := Arrondi(BaseTarifHT * TobLigne.GetValue('GL_QTEFACT'),V_PGI.OkDecV) ;;
  // BaseTarifHT:=ArrondirPrix(CodeArrondi,BaseTarifHT) ;  non géré actuellement dans l'article affaire
  TobDetAct.PutValue('ACT_PUVENTE', BaseTarifHT); TobDetAct.PutValue('ACT_PUVENTEDEV', BaseTarifHT);
  TobDetAct.PutValue('ACT_TOTVENTE', TotTarifHT); TobDetAct.PutValue('ACT_TOTVENTEDEV', TotTarifHT);
END;

Procedure CalculPRLigneAct (TobLigne, TobDetAct, TobArt : TOB);
Var PUPR,PUPRCharge, TotPRCharge : Double;
BEGIN
PUPR := TobDetAct.GetValue('ACT_PUPR');
// PUPRCharge := TobArt.GetValue('GA_DPR'); pb si même article utilisé x fois (article générique)
PUPRCharge := CalculPrixRevient(TobLigne.GetValue('GL_CODEARTICLE'), TobLigne.GetValue('GL_TIERS'),
                                TobLigne.GetValue('GL_DEPOT'),PUPR);
if PUPRCharge = 0 then PUPRCharge := PUPR;
TotPRCharge := Arrondi(PUPRCharge * TobDetAct.GetValue('ACT_QTE'),V_PGI.OkDecV) ;;
TobDetAct.PutValue('ACT_PUPRCHARGE', PUPRCharge);
TobDetAct.PutValue('ACT_TOTPRCHARGE', TOTPRCharge);
END;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 19/01/2001
Modifié le ... : 19/01/2001
Description .. : maj de la zone tenue stock / traitement de cette ligne en 
Suite ........ : activité. 
Suite ........ : Principe : sur pièces d'achats si la ligne passe en activité 
Suite ........ : (association à une affaire), elle n'est pas intégrée au stock. 
Mots clefs ... : 
*****************************************************************}
Procedure ImpactActiviteLigne(TOBPiece : TOB; GereActivite : Boolean);
Var i : integer;
    TobL : TOB;
BEGIN
//if not(ctxAffaire in V_PGI.PGIContexte) and not(ctxGCAFF in V_PGI.PGIContexte) then Exit;
// gm 04/04/03 pbm , en GC, si saisie piece achat avec code affaire, le stock n'était plu
// géré
// Je laisse comme ca pour affaire !!!
if not(ctxAffaire in V_PGI.PGIContexte)  then Exit;

if Not(GereActivite) then Exit;
if TobPiece=Nil then Exit;
For i:=0 To TobPiece.Detail.count-1 do
   BEGIN
   TobL := TobPiece.Detail[i];
   if (TobL.GetValue('GL_TENUESTOCK')<>'X') then Continue;
   if Not(EtudieLienligneToAct(TobL)) then Continue;
   // Ligne reprise en activité non passée en stock sur pièces d'achats
   if (GetInfoParPiece(TobL.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT')='ACH') then
      TobL.PutValue('GL_TENUESTOCK','-');
   END;
END;

end.
