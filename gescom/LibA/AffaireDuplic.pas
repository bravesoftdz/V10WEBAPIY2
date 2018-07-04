unit AffaireDuplic;

interface
uses HEnt1,HCtrls,AffaireUtil,EntGC,SysUtils, HMsgBox,
     FactGrp,UTOM,AGLInit,AffaireRegroupeUtil,Dicobtp,TiersUtil,
     UtilPGI, SAISUTIL,UTomAffaire, FactComm,AffaireModifPiece,Paramsoc,
{$IFDEF EAGLCLIENT}
     MainEAGL,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main,
{$ENDIF}
{$IFDEF BTP}
     CalcOleGenericBTP,
{$ENDIF}
     UTob,UtofProposToAffaire, FactTOB, FactUtil, FactAdresse, FactTiers ;


Type T_TypeDuplicAffaire =  (tdaDuplic,tdaAvenant,tdaAssistCreat,tdaProToAff);

function DuplicationAffaire (TypeDup : T_TypeDuplicAffaire; CodeAffRef, Arg : string; TobAffRef : TOB; bParle,RepFactaff,RepEtatAff : Boolean) : string;

Type T_DuplicAffaire = Class
       Constructor Create(TypeD : T_TypeDuplicAffaire; CodeAff, Arg : string; TobAff : TOB; bParle ,RepfActaff,RepEtatAff: Boolean);
       private
       TypeDup : T_TypeDuplicAffaire;
       CodeAffRef, CodeAffNew, CodeTiersNew, CodeTiers, StatutRef,StatutNew, AffNew0, AffNew1, AffNew2, AffNew3, AffNew4,
       AffDepartement, AffResponsable, NewNat, CodeAffMajbyPro :string;
       AffDateDebut, AffDateFin, AffDateDebGener, AffDateFinGener : TDateTime;
       MontantEcheDev : Double;
       Dev : RDEVISE;
       TOM_AFF : TOM;
       AvenantRemplace, RepriseSurAvenant,DateAffsurMoisClot,ExersurMoisClot,
       DechargeTobaffRef,EscompteTiers ,ModeRegleTiers,Parle,
       GarderCptPro, majAffRef,ActProToAff, RepriseEtatAff, RepriseFactAffRef, ProMajAffExist : Boolean;
       TobAffRef, TobAffTiersRef, TobFactaffRef,TobAffPieceRef, TobAdrRef,TobLienOleRef : TOB;
       TobAffNew,TobFactaffNew,TobAffPieceNew, TobAdrNew,TobAffTiersNew, TobTiersNew,TOBPieceRef ,TobLienOleNew : TOB;
       NumError : integer;


       Function  DuplicationAffaire : string;
       Function  VerifTiersValide : Boolean;
       Function  CodificationNewAff : Boolean;
       Function  DuplicTobRefEnNew  : Boolean;
       Function  MajClientsurNewAff : Boolean;
       Function  MajCleAffairesurNewAff :Boolean;
       Function  MajAffairesurNewAff :Boolean;
       Function  MajAffairesurAffRef : Boolean;
       Procedure MajAffSpecifAvenant;
       // Traitement des dates
       Procedure TraiteDateNewAff;
       Procedure RecalculDateAff;
       Procedure AugmenteDateAff (Nb : integer; unTps : string);

       Function  ChargeLesTobAff : Boolean;
       Procedure ToutAllouer ;
       Procedure ToutLiberer ;
       Procedure GetDevise;
       Procedure GetNaturePiece;
       // Validation
       Function ValideDuplication : Boolean;
       Procedure ValideActivite;
       Function  ValideAdresses : Boolean;
       Procedure CalculFactAff;
       // spécif proposition => Affaire sur affaire existante
       end;

// libellés des messages de la duplication d'affaire
Const	MsgCreat: array[1..9] of string 	= (
          {1}        'Code affaire déja utilisé pour ce client',
          {2}        'Nouveau code affaire incorrect',
          {3}        'Problème rencontré lors de la validation',
          {4}        'Validation de l''affaire impossible',
          {5}        'pièce d''origine non valide',
          {6}        'Duplication d''affaires',
          {7}        'La création a échouée',
          {8}        'Le client a un état comptable rouge',
          {9}        'Le Client est fermé'
          );

Procedure InitAppreciation (TobDet : TOB);
Procedure RecalculTotAffaire (RecupTotaux : string; TobAffNew : TOB);
Procedure MajTotAffaireParLesDEV (TobAff : TOB; bTenueEuro:boolean);
Function DuplicationTache(pStAffaireSource, pStAffaireDest : String) : Boolean;

implementation
{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 18/07/2001
Modifié le ... : 18/07/2001
Description .. : Fonction de duplication d'affaire;
Mots clefs ... : DUPLICATION AFFAIRE
*****************************************************************}
function DuplicationAffaire (TypeDup : T_TypeDuplicAffaire; CodeAffRef, Arg : string; TobAffRef : TOB; bParle,RepFactAff ,RepEtatAff: Boolean) : string;
Var DupAff : T_DuplicAffaire;
begin
Result := '';
DupAff:=T_DuplicAffaire.Create(TypeDup,CodeAffRef,Arg,TobAffRef,bParle,RepFactAff, RepetatAff);
Result := DupAff.DuplicationAffaire;
DupAff.Free ;
end;

Constructor T_DuplicAffaire.Create ( TypeD : T_TypeDuplicAffaire; CodeAff, Arg : string; TobAff : TOB; bParle,RepFactaff ,RepEtatAff: Boolean );
Var
   Tmp,Champ, Valeur : string;
   X : integer;
begin
// initialisation
AvenantRemplace := False;  RepriseSurAvenant:=False; DateAffsurMoisClot:=False;
EscompteTiers := False;  ModeRegleTiers :=False;
GarderCptPro := False;  ActProToAff := False;
MajAffRef := False;
ProMajAffExist := False;
ExersurMoisClot := False;
AffDepartement := ''; AffResponsable := '';
AffDateDebut := iDate1900; AffDateFin := iDate2099;
AffDateDebGener := iDate1900; AffDateFinGener := iDate2099;
MontantEcheDev := 0;

// Recup param
CodeAffRef:=CodeAff;
TOBAffRef:=TobAff;
TypeDup:=TypeD;
Parle := bParle;
RepriseFactAffRef := ( (TypeDup = tdaDuplic) or ((TypeDup = tdaProToAff) And (ProMajAffExist)) ); // tdaAvenant ??
If (RepFactaff) and (typeDup=tdaAssistCreat) then RepriseFactAffRef:=true;  // cas asisstant cr"ation ave coché de conserver info personnalisée de fatcaff
RepriseEtatAff :=RepEtatAff; // si reprise état affaire de la fiche d'origine
Tmp:=(Trim(ReadTokenSt(Arg)));
While (Tmp <> '') do
    begin
    X:=pos(':',Tmp);
    if x<>0 then begin Champ:=copy(Tmp,1,X-1); Valeur:=Copy (Tmp,X+1,length(Tmp)-X); end
            else Champ := Tmp;
    // Recup des arguments ...
    if Champ = 'AFF_TIERS'          then CodeTiersNew := Valeur;
    if Champ = 'AFF_RESPONSABLE'    then AffResponsable := Valeur;
    if Champ = 'AFF_DEPARTEMENT'    then AffDepartement := Valeur;
    if Champ = 'AFF_DATEDEBUT'      then AffDateDebut := strToDate(Valeur);
    if Champ = 'AFF_DATEFIN'        then AffDateFin := strToDate(Valeur);
    if Champ = 'AFF_DATEDEBGENER'   then AffDateDebGener := strToDate(Valeur);
    if Champ = 'AFF_DATEFINGENER'   then AffDateFinGener := strToDate(Valeur);
    if Champ = 'DATEAFFSURMOISCLOT' then DateAffsurMoisClot := (Valeur ='X');
    if Champ = 'EXERSURMOISCLOT'    then ExersurMoisClot := (Valeur ='X');
    if Champ = 'AFF_MONTANTECHEDEV' then MontantEcheDev := strtoFloat(Valeur);
    if Champ = 'ESCOMPTETIERS'      then EscompteTiers := (Valeur ='X');
    if Champ = 'MODEREGLETIERS'     then ModeRegleTiers :=(Valeur ='X');
      // spécif avenant ...
    if Champ = 'AVENANTREMPLACE'    then AvenantRemplace := (Valeur ='X');
    if Champ = 'REPRISESURAVENANT'  then RepriseSurAvenant := (Valeur ='X');
      // spécif proposition en affaire ...
    if Champ = 'GARDERCPTPRO'       then GarderCptPro :=(Valeur ='X');
    if Champ = 'ACTPROTOAFF'        then ActProToAff :=(Valeur ='X');
    if Champ = 'AFFAIREMAJBYPRO'    then begin CodeAffMajbyPro:=Valeur; ProMajAffExist:=true; end;
    Tmp:=(Trim(ReadTokenSt(Arg)));
    end;
end;

Function T_DuplicAffaire.DuplicationAffaire : string;
Var bTraiteok : Boolean;
begin
Result:= '';
DechargeTobAffRef := (TobAffRef=nil);
ToutAllouer;
bTraiteOk := ChargeLesTobAff ; // Chargement de la tob Affaire

// *** cas particulier : Maj d'une pièce existante ***
if ((TypeDup = tdaProToAff) And (ProMajAffExist)) then
   begin
   bTraiteOk := RemplirTOBAffaire(CodeAffMajbyPro,TobAffNew);
   end
else // *** Cas Général : charge des tob + duplication ***
   begin
   // Traitement du nouveau tiers
   if bTraiteOk = True then
      bTraiteOk := (RemplirTOBTiers (TOBTiersNew,CodeTiers,'',False) = trtOk);
   if bTraiteOk = True then bTraiteOk := VerifTiersValide;
   // Traitement nouvelle affaire
   if bTraiteOk = True then bTraiteOk := CodificationNewAff;
   if bTraiteOk = True then bTraiteOk := DuplicTobRefEnNew;

   if (bTraiteOk = True) and (CodeTiersNew <> '') then bTraiteOk := MajClientsurNewAff;

   if bTraiteOk = True then bTraiteOk := MajCleAffairesurNewAff;
   if bTraiteOk = True then bTraiteOk := MajAffairesurNewAff;
   end;

if bTraiteOk = True then bTraiteOk := MajAffairesurAffRef;
   // mcd 22/01/02 transac non imbriqués io := Transactions(ValideDuplication,1);
if bTraiteOk = True then   BtraiteOk:=ValideDuplication;
if BtraiteOk = False then
   begin
   Case V_PGI.IOError of
      oeUnknown : Numerror:=3;
      oeSaisie :  Numerror:=4;
      oePiece :   Numerror:=5;
      end;
   end;
if Parle then
   BEGIN
    if NumError > 0 then PgiBoxAf(CodeTiersNew+' '+CodeAffnew+' '+MsgCreat[7]+ ': '+MsgCreat[NumError],MsgCreat[6]);
   END;
ToutLiberer;
if bTraiteOk then Result := CodeAffNew
else result:='';
end;

Procedure T_DuplicAffaire.ToutAllouer ;
begin
if TobAffRef = nil then
   TobAffRef   := TOB.Create ('AFFAIRE',Nil,-1);
TobAffTiersRef := TOB.Create ('',Nil,-1);
TobFactaffRef  := TOB.Create ('',Nil,-1);
TobAffPieceRef := TOB.Create ('',Nil,-1);
TobAdrRef      := TOB.Create ('',Nil,-1);
TobLienOleRef  := TOB.Create ('',Nil,-1);   //mcd 10/02/03
TobPieceRef    := TOB.Create ('PIECE',Nil,-1);
// modif BTP
AddLesSupEntete (TOBPieceRef);
// ---

TobAffNew      := TOB.Create ('AFFAIRE',Nil,-1);
TobAffTiersNew := TOB.Create ('',Nil,-1);
TobFactaffNew  := TOB.Create ('',Nil,-1);
TobAffPieceNew := TOB.Create ('',Nil,-1);
TobLienOleNew  := TOB.Create ('',Nil,-1);   //mcd 10/02/03
TobAdrNew      := TOB.Create ('',Nil,-1);
TobTiersNew    := TOB.Create ('TIERS',Nil,-1);  TOBTiersNew.AddChampSup('RIB',False) ;

TOM_AFF := CreateTom('AFFAIRE',Nil,False, False);
end;

Procedure T_DuplicAffaire.ToutLiberer ;
begin
if (DechargeTobaffRef) and (TobAffRef <> Nil) then TobAffRef.Free;
if TobAffTiersRef <> Nil   then TobAffTiersRef.Free;
if TobFactaffRef <> Nil    then TobFactaffRef.Free;
if TobAffPieceRef <> Nil   then TobAffPieceRef.Free;
if TobAdrRef <> Nil        then TobAdrRef.Free;
if TobLienOleRef <> Nil    then TobLienOleRef.Free;  //mcd 10/02/03
if TobPieceRef <> Nil      then TobPieceRef.Free;

if TobAffNew <> Nil        then TobAffNew.Free;
if TobAffTiersNew <> Nil   then TobAffTiersNew.Free;
if TobFactaffNew <> Nil    then TobFactaffNew.Free;
if TobAffPieceNew <> Nil   then TobAffPieceNew.Free;
if TobLienOleNew <> Nil    then TobLienOleNew.Free;  //mcd 10/02/03
if TobAdrNew <> Nil        then TobAdrNew.Free;
if TobTiersNew <> Nil      then TobTiersNew.Free;

if TOM_AFF <> Nil then TOM_AFF.Free;
end;

Procedure T_DuplicAffaire.GetDevise ;
begin
Dev.Code := TobAffNew.GetValue ('AFF_DEVISE');
if Dev.code = '' then DEV.Code:= V_PGI.DevisePivot;
GetInfosDevise(DEV);
end;

Procedure T_DuplicAffaire.GetNaturePiece;
begin
if Copy(CodeAffNew,1,1) = 'P' then StatutNew := 'PRO' else StatutNew := 'AFF';
if (TypeDup = tdaProToAFF) then StatutRef := 'PRO' else StatutRef := StatutNew;
if StatutNew='PRO' then NewNat :=VH_GC.AFNatProposition
                   else NewNat :=VH_GC.AFNatAffaire;
end;


Function T_DuplicAffaire.ChargeLesTobAff : boolean;
Var Q : TQuery;
    stSQL : String;
begin
Result := false;
// Tob Affaire
if (TobAffRef.GetValue('AFF_AFFAIRE')='') then
   begin
   if Not(RemplirTOBAffaire (CodeAffRef,TobAffRef)) then Exit;
   end
else
   if TobAffRef.GetValue('AFF_AFFAIRE') <> CodeAffRef then Exit;

Result := true;
if ((TypeDup = tdaProToAff) And (ProMajAffExist)) then Exit; //autres tob non chargées (maj du statut uniquement)

if (CodeTiersNew <> '') then CodeTiers := CodeTiersNew
                        else CodeTiers :=TobAffRef.GetValue('AFF_TIERS');

// Tob tiers sur Affaire ** duplication, il faut tout prendre ...
stSQL := 'SELECT * From AFFTIERS where AFT_AFFAIRE="' + CodeAffRef + '"';
Q:= OpenSQL(stSQL,True,-1,'',true);
if Not Q.EOF then TobAffTiersRef.LoadDetailDB ('AFFTIERS','','',Q,False,False);
Ferme(Q);

// Tob affairePiece  ** duplication, il faut tout prendre ...
stSQL := 'SELECT * From AFFAIREPIECE where API_AFFAIRE="' + CodeAffRef + '"';
Q:= OpenSQL(stSQL,True,-1,'',true);
if Not Q.EOF then TobAffPieceRef.LoadDetailDB ('AFFAIREPIECE','','',Q,False,False);
Ferme(Q);



// Tob des FactAff
// (chargée dans certains cas uniquement, sinon recalculé en auto par update de la tom)

if RepriseFactAffRef then
   begin
        // 562 mcd 12/01/01 ajout test tdaduplic
        // PL le 06/03/02 : INDEX 3  ** duplication, il faut tout prendre ...
    //mcd 16/09/03 ajout TdaAssistcreat : si on est dans le cas d'assistant on reprend toutes les eche même facturée
    // reste le Cas tdaProToAff ou on ne reprend pas les ech facturée (?? est normal ???)
  if (TypeDup = tdaDuplic) or (typeDup=tdaAssistCreat) then  stSQL := 'SELECT * From FACTAFF where AFA_TYPECHE="NOR" AND AFA_AFFAIRE="' + CodeAffRef + '"'
   else    stSQL := 'SELECT * From FACTAFF where AFA_TYPECHE="NOR" AND AFA_AFFAIRE="' + CodeAffRef + '" AND AFA_ECHEFACT="-"';
   Q:= OpenSQL(stSQL,True,-1,'',true);
   if Not Q.EOF then TobFactAffRef.LoadDetailDB ('FACTAFF','','',Q,False,False);
   Ferme(Q);
   end;

if (TypeDup <> tdaAssistCreat) or
   ((TypeDup = tdaAssistCreat) And (TobAffRef.GetValue('AFF_TIERS')=CodeTiers)) then
   begin
   // PL le 06/03/02 : INDEX 2  ** duplication, il faut tout prendre ...
   stSQL := 'SELECT * From ADRESSES where ((ADR_TYPEADRESSE="AFA") OR (ADR_TYPEADRESSE="INT")) AND ADR_REFCODE="' + CodeAffRef + '"';
   Q:= OpenSQL(stSQL,True,-1,'',true);
   if Not Q.EOF then TobAdrRef.LoadDetailDB ('ADRESSES','','',Q,False,False);
   Ferme(Q);
   end;
if (TypeDup = tdaDuplic) or  (TypeDup = tdaAssistCreat) then  // mcd 10/02/03
  begin
   stSQL := 'SELECT * From LIENSOLE where LO_TABLEBLOB="AFF" AND LO_IDENTIFIANT="'+ CodeAffRef + '"';
   Q:= OpenSQL(stSQL,True,-1,'',true);
   if Not Q.EOF then TobLienOleRef.LoadDetailDB ('LIENSOLE','','',Q,False,False);
   Ferme(Q);
  end;
end;

Function T_DuplicAffaire.CodificationNewAff  : Boolean;
Var
   bSaisieManuelle, bRecodif : Boolean;
   ArgType, Aff0 : string;
begin
bSaisieManuelle := False;
Result := True;
bRecodif := GetParamSoc ('SO_AFFRECODIFICATION');
// Recodification du code de la nouvelle affaire
AffNew0 := TobAffRef.GetValue('AFF_AFFAIRE0');
AffNew1 := TobAffRef.GetValue('AFF_AFFAIRE1');
AffNew2 := TobAffRef.GetValue('AFF_AFFAIRE2');
AffNew3 := TobAffRef.GetValue('AFF_AFFAIRE3');
AffNew4 := TobAffRef.GetValue('AFF_AVENANT');
if (TypeDup = tdaduplic) then AffNew2 := CodificationNewExercice (AffNew2,true);


// Si Scot Test d'existence de la mission exercice avant augmentation du compteur
   // mcd 18/09/01 if (ctxScot in V_PGI.PGIContexte) And (TypeDup <> tdaAvenant) then
if (GetParamSoc('SO_AFFORMATEXER') <>'AUC')  And (TypeDup <> tdaAvenant)  then
   BEGIN
   if (typeDup = tdaProToAff) then Aff0 := 'A'
                              else Aff0 := AffNew0;
   if ChercheMissionExercice (Aff0,AffNew1,AffNew2, AffNew4,CodeTiers) then
      BEGIN NumError := 1; Result:=False; Exit; END;
   END;

// ************* En Création d'avenant ************************
if (TypeDup = tdaAvenant) then
   begin
   CodeAffNew := GetCodeAffAvSuivant (CodeAffRef);
   AffNew4 := Copy(CodeAffNew,16,2);
   end;
   // En création d'avenant, le code généré est forcement uniquement - controle inutile

// ********************* proposition to affaire *************************
if (TypeDup = tdaProToAff) then
   begin
   CodeAffNew := PropositionToAffaire(CodeAffRef, AffNew0, AffNew1, AffNew2, AffNew3, AffNew4,GarderCptPro,Not(bRecodif));
   // PA le 24/08/2001 - si pas de recodif update cpt affaire sinon pas d'update car on passe par dechargecleaffaire si valide le cptQ
   end;

// ********************* Duplication Exercice N en N+1 ************************
if (TypeDup = tdaduplic) then
   begin
   CodeAffNew :=CodeAffaireRegroupe(AffNew0, AffNew1, AffNew2, AffNew3, AffNew4, taCreat, True,false,True);
   end;
// ********************* Assistant de création d'affaires ********************
if (TypeDup= tdaAssistCreat) then
   begin
   CodeAffNew :=CodeAffaireRegroupe(AffNew0, AffNew1, AffNew2, AffNew3, AffNew4, taCreat, True,false,Not(bRecodif));
   // PA le 24/08/2001 - si pas de recodif update cpt affaire sinon pas d'update car on passe par dechargecleaffaire si valide le cpt
   end;

// ********************* Test d'existance de la nouvelle affaire **************
   // mcd 18/09/01 if (Not(ctxScot in V_PGI.PGIContexte)) and
if (GetParamSoc('SO_AFFORMATEXER') ='AUC') And
       ((TypeDup = tdaAssistCreat) or (TypeDup = tdaProToAff)) then
   begin
   while (CodeAffNew <> '') and (ExisteAffaire(CodeAffNew,'') or (bRecodif)) do
      begin
      if (TypeDup = tdaAssistCreat) then ArgType:= 'TYPE:AFFAIREMODELE;'
                                    else ArgType:= '';
      if bRecodif then ArgType := ArgType + 'RECODIFICATION;';
      ArgType := ArgType + 'AFFAIREENCREATION;'; // Pour UtofafBaseCodeAffaire + chargeCle du code Aff
      ArgType := ArgType + 'AFFAIRE:'+ CodeAffNew + ';';
      ArgType := ArgType + 'PROPOSITION:'+ CodeAffRef + ';';
      ArgType := ArgType + 'TIERS:'+ CodeTiers + ';';
      ArgType := ArgType + 'LIBTIERS:'+ TobTiersNew.getvalue('T_LIBELLE') + ';';
      CodeAffNew:=AFLanceFiche_ProposToAffaire(ArgType);
      bSaisieManuelle := true;
      bRecodif := False;
      end;
  end;

if (CodeAffNew = '') then BEGIN NumError := 2; Result:=False; Exit; END;
if bSaisieManuelle then
   {$IFDEF BTP}
   BTPCodeAffaireDecoupe(CodeAffNew,AffNew0, AffNew1, AffNew2, AffNew3, AffNew4, taCreat, false);
   {$ELSE}
   CodeAffaireDecoupe(CodeAffNew,AffNew0, AffNew1, AffNew2, AffNew3, AffNew4, taCreat, false);
   {$ENDIF}
   GetNaturePiece;
end;


Function T_DuplicAffaire.DuplicTobRefEnNew  : Boolean;
begin
TobAffNew.Dupliquer (TobAffRef,False,True,False);
TobAffTiersNew.Dupliquer (TobAffTiersRef,True,True,False);
TobFactAffNew.Dupliquer (TobFactAffRef,True,True,False);
TobAffPieceNew.Dupliquer (TobAffPieceRef,True,True,False);
TobAdrNew.Dupliquer (TobAdrRef,True,True,False);
TobLienOleNew.Dupliquer (TobLienOleRef,True,True,False); //mcd 10/02/03
Result := True;
end;




Function T_DuplicAffaire.MajClientsurNewAff : Boolean;
BEGIN
Result := True;
TobAffNew.PutValue('AFF_TIERS', TobTiersNew.GetValue ('T_TIERS'));
TobAffNew.PutValue('AFF_CONTACT',''); // pas sur le même client ...
END;

Function T_DuplicAffaire.VerifTiersValide : Boolean;
Var EtatRisque : string;
BEGIN
Result := False;
// Tiers Fermé
if TobTiersNew.GetValue('T_FERME')='X' then begin NumError := 9; Exit; end;

// Risque client
EtatRisque :=GetEtatRisqueClient (CodeTiers);
if EtatRisque = 'R' then begin NumError := 8; Exit; end;

Result := True;
END;

Function T_DuplicAffaire.MajCleAffairesurNewAff :Boolean;
Var
   i : integer;
   TobDet : TOB;
begin
Result := True;
// TobAffaire
TobAffNew.PutValue('AFF_AFFAIRE' ,CodeAffNew);
TobAffNew.PutValue('AFF_AFFAIRE0',AffNew0);
TobAffNew.PutValue('AFF_AFFAIRE1',AffNew1);
TobAffNew.PutValue('AFF_AFFAIRE2',AffNew2);
TobAffNew.PutValue('AFF_AFFAIRE3',AffNew3);
TobAffNew.PutValue('AFF_AVENANT' ,AffNew4);
TobAffNew.PutValue('AFF_STATUTAFFAIRE', StatutNew);

// TobAfftiers
for i := 0 to TobAffTiersNew.detail.count-1 do
   begin
   TobDet := TobAffTiersNew.detail[i];
   TobDet.PutValue('AFT_AFFAIRE', TobAffNew.GetValue('AFF_AFFAIRE'));
   end;
// Factaff
for i := 0 to TobFactaffNew.detail.count-1 do
   begin
   TobDet := TobFactaffNew.detail[i];
   InitAppreciation (TobDet);
   TobDet.PutValue('AFA_AFFAIRE', TobAffNew.GetValue('AFF_AFFAIRE'));
   TobDet.PutValue('AFA_TIERS', TobAffNew.GetValue('AFF_TIERS')); //mcd 24/09/03
   end;
// AffairePiece
for i := 0 to TobAffPieceNew.detail.count-1 do
   begin
   TobDet := TobAffPieceNew.detail[i];
   InitAppreciation (TobDet);
   TobDet.PutValue('API_AFFAIRE', TobAffNew.GetValue('AFF_AFFAIRE'));
   end;
// Adresse
for i := 0 to TobAdrNew.detail.count-1 do
   begin
   TobDet := TobAdrNew.detail[i];
   TobDet.PutValue('ADR_REFCODE', TobAffNew.GetValue('AFF_AFFAIRE'));
   end;
// ALeinsOle              mcd 10/02/03
for i := 0 to TobLienOleNew.detail.count-1 do
   begin
   TobDet := TobLienOleNew.detail[i];
   TobDet.PutValue('LO_IDENTIFIANT', TobAffNew.GetValue('AFF_AFFAIRE'));
   end;

end;

Function T_DuplicAffaire.MajAffairesurNewAff :Boolean;
Var
    TypeGener : String;
    MtDev,MtPivot : double;
BEGIN
Result := True;
GetDevise;
// Zones générales modifiées
TobAffNew.PutValue('AFF_CREATEUR', V_PGI.User);
TobAffNew.PutValue('AFF_UTILISATEUR', V_PGI.User);
if (TypeDup <> tdaDuplic) then TobAffNew.PutValue('AFF_MODELE','-'); //mcd 17/10/02 su duplic N en N+1 il faur conserver l'info , si assistant à supprimer
TobAffNew.PutValue('AFF_CREERPAR','GEN');
TobAffNew.PutValue('AFF_DATECREATION', V_PGI.DateEntree);
TobAffNew.PutValue('AFF_DATEMODIF', V_PGI.DateEntree);
TobAffNew.PutValue('AFF_NUMDERGENER','');

// mcd 10/10/2002 ajout des nouveaux cahmps qui doivent aussi être initialisé
TobAffNew.PutValue('AFF_APPRECIATION','');
TobAffNew.PutValue('AFF_NUMSITUATION',0);
TobAffNew.PutValue('AFF_DATECUTOFF',idate1900);
TobAffNew.PutValue('AFF_DATESITUATION',idate1900);
   //mcd 17/09/03 ajout test sur rerpise état origine ou pas (peut être vrai pour assistant créationna ffaire uniquement)
If not (RepriseEtatAff) then TobAffNew.PutValue ('AFF_ETATAFFAIRE','ENC'); // MCD 18/12/02 lors de la duplic on met la même valeur qu'en création d'affaire



// traitement affaire de référence + affaire complète
if (TypeDup = tdaAvenant) then // l'affaire de ref de l'avenant idem aff.
   begin
   if TobAffRef.GetValue('AFF_AFFAIREREF') = TobAffRef.GetValue('AFF_AFFAIRE') then
      TobAffNew.PutValue('AFF_AFFAIREREF',TobAffNew.GetValue('AFF_AFFAIRE'))
   else
      TobAffNew.PutValue('AFF_AFFAIREREF',TobAffRef.GetValue('AFF_AFFAIREREF'));
   end
else        // Affaire indépendante en duplication
   begin
   TobAffNew.PutValue('AFF_AFFAIREREF',TobAffNew.GetValue('AFF_AFFAIRE'));
   TobAffNew.PutValue('AFF_ISAFFAIREREF','-');
   end;

// Récup des zones paramétrables
if AffResponsable <> '' then  TobAffNew.PutValue('AFF_RESPONSABLE',AffResponsable);
if AffDepartement <> '' then  TobAffNew.PutValue('AFF_DEPARTEMENT',AffDepartement);

// Traitement des dates
TraiteDateNewAff;

// ******************** maj spécif à l'assistant de création ******************
if (TypeDup = tdaAssistCreat) then
   begin
   // Traitement du montant des écheances
   TypeGener := TobAffNew.GetValue('AFF_GENERAUTO');
   if TypeGener = 'FOR' then TobAffNew.PutValue('AFF_MONTANTECHEDEV',MontantEcheDev)
   else
   if TypeGener = 'POU' then TobAffNew.PutValue('AFF_POURCENTAGE',MontantEcheDev)
   else
   if TypeGener = 'POT' then    // ras toujours à 100 % non modifiable
   else
      TobAffNew.PutValue('AFF_MONTANTECHEDEV',MontantEcheDev);
   MtDev:= StrToFloat( TobAffNew.GetValue('AFF_MONTANTECHEDEV'));
   // C.B 18/06/2003 Suppression contrevaleur
   ConvertDevToPivot(DEV,MtDev,MtPivot);
   TobAffNew.PutValue('AFF_MONTANTECHE',Mtpivot);
   // Passage à l'état encours de la proposition
   if (StatutNew = 'PRO') And (TobAffRef.GetValue('AFF_ETATAFFAIRE')='ACC') then
      TobAffNew.PutValue('AFF_ETATAFFAIRE','ENC');
   end;
// ******************* Maj spécif à l'acceptation de propositions **************
if (TypeDup = tdaProToaff) then
   begin
   TobAffNew.PutValue('AFF_DATESIGNE', V_PGI.DateEntree);
   end;
// ********************Maj Specif à la création d'avenant **********************
if (TypeDup = tdaAvenant) then
   begin
   MajAffSpecifAvenant ;
   end;
END;


Procedure T_DuplicAffaire.TraiteDateNewAff;
begin
if (TypeDup = tdaAssistCreat) then
   begin
//   if (AffDateDebut=iDate1900) Or (AffDateFin=iDate2099) then Exit;
   if (AffDateDebut=iDate1900) and (AffDateFin=iDate2099) then Exit;    //gm 08/04/à2
   if (AffDateDebut > AffDateFin) then Exit;
   RecalculDateAff;
   end
else if (TypeDup = tdaduplic) then AugmenteDateAff (1, 'A');
end;

Procedure T_DuplicAffaire.RecalculDateAff;
Var
   DebutFact,FinFact,Liquidative,Garantie,Cloture,DebutExer,FinExer,
   DebutInter, FinInter : TDateTime;
   Exercice, part0,part1,part2,part3,ave : string;
   MoisCloture,MoisOld ,   annee,annee1 : Integer;
    QQ : TQuery ;
begin
     //17/10/01 pour OK si pas exercicie if ctxScot in V_PGI.PGIContexte then
if (VH_GC.AFFORMATEXER<>'AUC') then
   BEGIN
   Exercice := '';  // calcul basé sur le mois de cloture du client
   MoisCloture := TobTiersNew.GetValue ('T_MOISCLOTURE');
   if (MoisCloture < 1) or (Moiscloture >12) then MoisCloture := 12;
   if Not(DateAffsurMoisClot) then
      BEGIN
      DebutExer:=AffDateDebut; FinExer :=AffDateFin;
      END
   else
      BEGIN
      Exercice := TobAffNew.GetValue('AFF_AFFAIRE2');
      if (VH_GC.AFFORMATEXER = 'AM') then
         Exercice := copy(exercice,1,4) + format ('%2.2d',[Moiscloture])
                  // ajout 17/10/01 pas oK pour aa/aa si pas même mois cloture
      else if (VH_GC.AFFormatExer = 'AA') then   // **** année - année *****
           begin
           if CodeTiersNew <> TobAffRef.Getvalue('AFf_TIERS') then begin
              QQ:=OpenSQL('SELECT T_MOISCLOTURE FROM TIERS WHERE T_TIERS="'+TobAffREf.Getvalue('AFf_TIERS')+'"',TRUE,-1,'',true) ;
              if Not QQ.EOF then
                 begin
                 MoisOld:=QQ.Fields[0].AsInteger;
                 end else MoisOld:=12;
             Ferme(QQ);
             annee := StrtoInt (Copy(Exercice,1,4));
             annee1 := StrtoInt (Copy(Exercice,5,4));
             If MoisOld <> MoisCloture then
                If MOisOld=12 then inc (annee1)
                   else If MoisCloture =12 then annee1:=annee; // sinon, on laisse même année
             Exercice:=Format('%04d%4d',[annee,annee1]);
             end;
           end;
      if ExersurMoisClot then begin
         TobAffNew.PutValue('AFF_AFFAIRE2',Exercice);
         part0:= TobAffNew.GetValue('AFF_AFFAIRE0');
         part1:= TobAffNew.GetValue('AFF_AFFAIRE1');
         part2:= TobAffNew.GetValue('AFF_AFFAIRE2');
         part3:= TobAffNew.GetValue('AFF_AFFAIRE3');
         Ave:= TobAffNew.GetValue('AFF_AVENANT');
         TobaffNew.PutValue('AFF_AFFAIRE',RegroupePartiesAffaire(Part0,Part1,pArt2,part3,ave));   // mcd 17/10/01 il faut aussi remettre AFFAIRE Ok
         CodeAffNew:=TobAffNew.Getvalue('AFF_AFFAIRE');
         TobAffNew.PutValue('AFF_AFFAIREREF',TobAffNew.GetValue('AFF_AFFAIRE'));
         end;
      END;

   CalculDateExercice(VH_GC.AFFORMATEXER,MoisCloture,Exercice,DebutExer,Finexer,DebutInter,FinInter,DebutFact,FinFact,Liquidative,Garantie,Cloture,DateAffsurMoisClot);
   TobAffNew.PutValue('AFF_DATEDEBEXER',   DebutExer);
   TobAffNew.PutValue('AFF_DATEFINEXER',   finExer);
   TobAffNew.PutValue('AFF_DATEDEBUT',     DebutInter);
   TobAffNew.PutValue('AFF_DATEFIN',       finInter);
   TobAffNew.PutValue('AFF_DATEDEBGENER',  DebutFact);
   TobAffNew.PutValue('AFF_DATEFINGENER',  finFact);
   TobAffNew.PutValue('AFF_DATEFACTLIQUID',Liquidative);
   TobAffNew.PutValue('AFF_DATEGARANTIE',  Garantie);
   TobAffNew.PutValue('AFF_DATECLOTTECH',  Cloture);
   END
else  // En gestion d'affaires
   BEGIN
   TobAffNew.PutValue('AFF_DATEDEBUT',AffDateDebut);
   TobAffNew.PutValue('AFF_DATEFIN',AffDateFin);
   // Recalcul des dates en automatique ...
   CalculDateAffaire(AffDateDebut,AffDateFin,DebutFact,FinFact,Liquidative,Garantie,Cloture);
   // ??? PA garder les dates de facturation de l'affaire modèle ???
   if AffDateDebGener <> iDate1900 Then TobAffNew.PutValue ('AFF_DATEDEBGENER',  AffDateDebGener)
                                   Else TobAffNew.PutValue ('AFF_DATEDEBGENER',  DebutFact);
   if AffDateFinGener <> iDate2099 Then TobAffNew.PutValue ('AFF_DATEFINGENER',  AffDateFinGener)
                                   Else TobAffNew.PutValue('AFF_DATEFINGENER',  FinFact);
   TobAffNew.PutValue('AFF_DATECLOTTECH',  Cloture);
   TobAffNew.PutValue('AFF_DATEGARANTIE',  Garantie);
   END;
end;


Procedure T_DuplicAffaire.AugmenteDateAff (Nb : integer; unTps : string);
begin
if( TobAffNew.GetValue('AFF_DATEDEBEXER') <> idate1900)
    then TobAffNew.PutValue('AFF_DATEDEBEXER',PlusDate(TobAffNew.GetValue('AFF_DATEDEBEXER'),  Nb,UnTps));
if( TobAffNew.GetValue('AFF_DATEFINEXER') <> idate2099)
    then TobAffNew.PutValue('AFF_DATEFINEXER',PlusDate(TobAffNew.GetValue('AFF_DATEFINEXER'), Nb,UnTps));

if( TobAffNew.GetValue('AFF_DATEDEBUT') <> idate1900)
    then TobAffNew.PutValue('AFF_DATEDEBUT',PlusDate(TobAffNew.GetValue('AFF_DATEDEBUT'),     Nb,UnTps));
if( TobAffNew.GetValue('AFF_DATEFIN') <> idate2099)
    then TobAffNew.PutValue('AFF_DATEFIN',PlusDate(TobAffNew.GetValue('AFF_DATEFIN'),      Nb,UnTps));

if( TobAffNew.GetValue('AFF_DATEDEBGENER') <> idate1900)
    then TobAffNew.PutValue('AFF_DATEDEBGENER',PlusDate(TobAffNew.GetValue('AFF_DATEDEBGENER'),  Nb,UnTps));
if( TobAffNew.GetValue('AFF_DATEFINGENER') <> idate2099)
    then TobAffNew.PutValue('AFF_DATEFINGENER',PlusDate(TobAffNew.GetValue('AFF_DATEFINGENER'),  Nb,UnTps));

if( TobAffNew.GetValue('AFF_DATECLOTTECH') <> idate2099)
    then TobAffNew.PutValue('AFF_DATECLOTTECH',PlusDate(TobAffNew.GetValue('AFF_DATECLOTTECH'),Nb,UnTps));
// mcd 22/02/02 ... init à 2099 ...if( TobAffNew.GetValue('AFF_DATERESIL') <> idate1900)
if( TobAffNew.GetValue('AFF_DATERESIL') <> idate2099)
    then TobAffNew.PutValue('AFF_DATERESIL',PlusDate(TobAffNew.GetValue('AFF_DATERESIL'),Nb,UnTps));
if( TobAffNew.GetValue('AFF_DATELIMITE') <> idate2099)
    then TobAffNew.PutValue('AFF_DATELIMITE',PlusDate(TobAffNew.GetValue('AFF_DATERESIL'),Nb,UnTps));
if( TobAffNew.GetValue('AFF_DATEGARANTIE') <> idate2099)
    then TobAffNew.PutValue('AFF_DATEGARANTIE',PlusDate(TobAffNew.GetValue('AFF_DATEGARANTIE'),  Nb,UnTps));
if( TobAffNew.GetValue('AFF_DATEFACTLIQUID') <> idate2099)
    then TobAffNew.PutValue('AFF_DATEFACTLIQUID',PlusDate(TobAffNew.GetValue('AFF_DATEFACTLIQUID'),  Nb,UnTps));
// mcd 11/10/02 TobAffNew.PutValue('AFF_DATEFACTLIQUID',idate2099);
end;


Procedure T_DuplicAffaire.MajAffSpecifAvenant;
Var i : integer;
    TobDet : TOB;
begin

if Not(RepriseSurAvenant) then
   begin
   // Réinitialisation des montants
   TobAffNew.PutValue('AFF_MONTANTECHE',0); TobAffNew.PutValue('AFF_MONTANTECHEDEV',0);
   TobAffNew.PutValue('AFF_TOTALTTC',0);    TobAffNew.PutValue('AFF_TOTALTTCDEV',0);
   TobAffNew.PutValue('AFF_TOTALTAXE',0);   TobAffNew.PutValue('AFF_TOTALTAXEDEV',0);
   TobAffNew.PutValue('AFF_TOTALHT',0);     TobAffNew.PutValue('AFF_TOTALHTDEV',0);
   TobAffNew.PutValue('AFF_TOTALHTGLO',0);  TobAffNew.PutValue('AFF_TOTALHTGLODEV',0); 
   // raz des montants d'échéances
   for i := 0 to TobFactaffNew.Detail.count-1 do
      begin
      TobDet := TobFactaffNew.detail[i];
      TobDet.PutValue('AFA_MONTANTECHE', 0); TobDet.PutValue('AFA_MONTANTECHEDEV', 0); 
      end;
   end;
// A VOIR Date de début de génération des échéances de l'avenant  idem affaire ???

end;


Function  T_DuplicAffaire.ValideAdresses : Boolean;
Var i : integer;
    TobDet : TOB;
begin
result:=true;
for i := 0 to TobAdrNew.detail.count-1 do
   begin
   TobDet := TobAdrNew.detail[i];
   if not VH_GC.GCIfDefCEGID then
      CreerAdresse(TobDet, TobDet.GetValue('ADR_TYPEADRESSE')) ;
   end;
end;

Function T_DuplicAffaire.ValideDuplication:Boolean;
Var errModif : integer;
    bOK : Boolean;
    RecupTotaux : string;
BEGIN
// mcd 11/06/02 jamais mis  vrai , même si OK .. result:=false;
result:=True;
RecupTotaux:='';
BOk := true;
if Not((TypeDup = tdaProToAff) And (ProMajAffExist)) then
   begin
   TobAffNew.SetAllModifie(true);
   TobAffTiersNew.SetAllModifie(true);
   TobFactAffNew.SetAllModifie(true);
   end;

   /////// MAJ de la piece ////////////////
if (TypeDup  = tdaDuplic)   or  (TypeDup = tdaAssistCreat) or
   ((TypeDup = tdaProToAff) And Not(ProMajAffExist)) or
   ((TypeDup = tdaAvenant)  And (RepriseSurAvenant)) then
   begin
   if Not(GetTobPieceAffaire(CodeAffref ,StatutRef,TobPieceRef)) then
      begin V_PGI.IoError:=oePiece ;exit; end;
   DupliqueLaPiece ( TOBPieceRef, TOBAffNew, NewNat, CodeTiersNew,False,EscompteTiers,ModeRegleTiers,V_PGI.DateEntree);
   end
else // *** maj d'une pièce existante ***
if ((TypeDup = tdaProToAff) And (ProMajAffExist)) then
   begin
   if Not(GetTobPieceAffaire(CodeAffref ,StatutRef,TobPieceRef)) then
      begin V_PGI.IoError:=oePiece ;exit; end;
   errModif := ModifieLaPieceAff (tmpAjoutLigne, CodeAffMajbyPro, TOBPieceRef,False,RecupTotaux);
   if errModif <> 0 then
      V_PGI.IoError:=oeUnknown
   else
      RecalculTotAffaire(RecupTotaux,TobAffNew);
   end;

////// MAJ de l'affaire /////////////
TobAffNew.AddChampSupValeur ('CreationAuto','X');
if Not(RepriseFactAffRef) then TobAffNew.AddChampSupValeur ('CalculEcheances','X');
If Not(TOM_AFF.VerifTOB (TobAffNew)) then begin V_PGI.IoError:=oeSaisie ;Result:=false; exit; end;
if ((TypeDup = tdaProToAff) And (ProMajAffExist)) then
   begin //**** Cas spécifique prop => Affaire sur affaire existante ****
   if (V_PGI.IoError=oeOk) then
      bOK := TobAffNew.UpdateDBTable;
   if not(bOK) then V_PGI.IoError:=oeUnknown ;
   end
else  // ****** Cas général de duplication *****
   begin
   // création des écheances
   CalculFactAff ; // Sauf cas particulier fait en auto par le veriftob (update de la tom)
   if V_PGI.IoError=oeOk then
      if Not(TobAffNew.InsertDB(Nil,False)) then begin V_PGI.IoError:=oeUnknown ;Result:=False;exit;end; //mcd 26/02/02 ajout arrêt si erreur
      // mcd 25/11/02 if Not(TobAffNew.InsertDBTable(Nil,False)) then begin V_PGI.IoError:=oeUnknown ;Result:=False;exit;end; //mcd 26/02/02 ajout arrêt si erreur
   if TobAffTiersNew.Detail.count<>0 then TobAffTiersNew.InsertDB(Nil,False);
   if TobFactAffNew.Detail.count <>0 then TobFactAffNew.InsertDB (Nil,False);
   if TobAffPieceNew.Detail.count <>0 then TobAffPieceNew.InsertDB (Nil,False);
   if TobLienOleNew.Detail.count <>0 then TobLienOleNew.InsertDB (Nil,False); //mcd 10/02/03
   ValideAdresses;
   // transfert de l'activité de l'affaire de référence si nécessaire
   if ActProToAff then ValideActivite;
   end;

// Maj de l'affaire de référence si modif à valider ( ex: proposition acceptée)
if MajAffRef then TobAffRef.UpdateDBTable;
if V_PGI.IoError<>oeOk then    result:=false;
end;

Procedure T_DuplicAffaire.CalculFactAff;
Var str1,str2,str3,str4,str5,str6,str7 : string;
    i ,ii,j,jj,annee: integer;
    TobDet : TOB;
BEGIN
if (TypeDup = tdaDuplic) then
   begin
   // dans le cas de duplication, on incrémente les dates
   if (TobFactAffNew <> Nil) then
   begin
   for i := 0 to TobFactaffNew.detail.count-1 do
       BEGIN
       TobDet := TobFactaffNew.detail[i];
       InitAppreciation (TobDet);
       if (Tobdet.GetValue('AFA_DATEECHE') <>idate1900) then TobDet.PutValue('AFA_DATEECHE',PlusDate(Tobdet.GetValue('AFA_DATEECHE'),  1,'A'));
          // si année dans le texte, on incrémente....)
       ii := pos('/',Tobdet.GetValue('AFA_LIBELLEECHE'));
       if ii <>0 then begin
          str2:=  Tobdet.GetValue('AFA_LIBELLEECHE');
          jj:= strlen(Pchar(str2));
          str2:=Copy(Tobdet.GetValue('AFA_LIBELLEECHE'),ii+1,jj);
          J := pos('/',str2);
          str2:='';
          if (j <>0) then begin
             if (j <=3) then ii:=II+J;   // pour cas ou MM/AAAA puis MM/AAAA . sera OK si JJ/MM/AAAA
             end;
          str1:=  Copy(Tobdet.GetValue('AFA_LIBELLEECHE'),1,ii) ;
          j:=JJ;
          if (ii+5 < jj) then begin
                     // on regarde si il y a une suite de libellé
                   str6:=Copy(Tobdet.GetValue('AFA_LIBELLEECHE'),ii+1,jj-ii+1);
                   J := pos(' ',str6);
                   if (j<>0) then J:=J+II else J:=jj;
                      // puis on regarde si le la date trouvé précédemment est bien sur 4c
                   if (J = ii+5) then    str2:=Copy(Tobdet.GetValue('AFA_LIBELLEECHE'),ii+5,jj-(ii+4))
                   else str2:=Copy(Tobdet.GetValue('AFA_LIBELLEECHE'),J+1,jj-J);
                   end;
                    //mcd 16/09/03 ajout test AAA suite pb en chez un lcient qui avait ce cas ..
          if ((ii+4) <=J) and (Copy(Tobdet.GetValue('AFA_LIBELLEECHE'),ii+1,4) <>'AAAA') then
            begin
            annee := StrToInt(Copy(Tobdet.GetValue('AFA_LIBELLEECHE'),ii+1,4)) ;
            if (annee > 1901) then begin
                      //On ne le fait que si date correcte
               inc(annee);
               str3:=format ('%4.4d',[annee]);
               TobDet.putvalue('AFA_LIBELLEECHE',str1+str3+str2);
               end;
            end
            else begin
               str3:=Copy(Tobdet.GetValue('AFA_LIBELLEECHE'),ii+1,j- ii);
               TobDet.putvalue('AFA_LIBELLEECHE',str1+str3+str2);
              end;
             // on recommence le traitement si 2 dates dans le libellés
         ii := pos('/',str2);
         if ii <>0 then begin
            str5:=  str2;
            jj:= strlen(Pchar(str5));
            str5:=Copy(str2,ii+1,jj);
            J := pos('/',str5);
            str5:='';
            if (j <>0) then begin
               ii:=II+J;
               end;
            str4:=  Copy(str2,1,ii) ;
            jj:= strlen(Pchar(str2));
            j:=JJ;
            if (ii+5 < jj) then begin
                       // on regarde si il y a une suite de libellé
                     str7:=Copy(str2,ii+1,jj-ii+1);
                     J := pos(' ',str7);
                     if (j<>0) then J:=J+II else J:=jj;
                        // puis on regarde si le la date trouvé précédemment est bien sur 4c
                     if (J = ii+5) then    str5:=Copy(str2,ii+5,jj-(ii+4))
                     else str5:=Copy(str2,J+1,jj-J);
                     end;
            if (ii+4) <=J then begin
              annee := StrToInt(Copy(str2,ii+1,4));
              if (annee > 1901) then begin
                        //On ne le fait que si date correcte
                 inc(annee);
                 str6:=format ('%4.4d',[annee]);
                 TobDet.putvalue('AFA_LIBELLEECHE',str1+str3+ str4 + str6 + str5);
                 end;
            end
         else begin
               str6:=Copy(str2,ii+1,j-ii);
               TobDet.putvalue('AFA_LIBELLEECHE',str1+str3+ str4 + str6 + str5);
              end;
            end;

          end;
       end;
     end;
   end;
END;


Function T_DuplicAffaire.MajAffairesurAffRef : Boolean;
begin
Result := True;
if (TypeDup = tdaProToAff) then // **** Modif sur la proposition acceptée ******
   begin
   MajAffRef := True;
   TobAffRef.PutValue('AFF_ETATAFFAIRE', 'ACC');
   TobAffRef.PutValue('AFF_DATESIGNE', V_PGI.DateEntree);
   end;
// Maj de l'état de l'affaire de référence si elle est remplacée par l'avenant
if (TypeDup = tdaAvenant) And (AvenantRemplace) then
   begin
   MajAffRemplaceeParAvenant (TobAffRef);
   MajAffRef := True;
   end;
end;


Procedure T_DuplicAffaire.ValideActivite;
Var stMaj : String;
begin
stmaj := '';
if (TypeDup = tdaProToAff) then stMaj  := ', ACT_AVANTVENTE="X"';

ExecuteSQL('UPDATE ACTIVITE SET ACT_AFFAIRE="'+ CodeAffNew +'", ACT_AFFAIRE0="'+ AffNew0
           +'", ACT_AFFAIRE1="'+ AffNew1
           +'", ACT_AFFAIRE2="'+ AffNew2
           +'", ACT_AFFAIRE3="'+ AffNew3
           +'", ACT_AVENANT="' + AffNew4 + '"'
           + stMaj
           +' WHERE ACT_TYPEACTIVITE="REA" AND ACT_AFFAIRE="' + CodeAffRef +'" ');
end;
                        


Procedure RecalculTotAffaire (RecupTotaux : string; TobAffNew : TOB);
Var bBaseEch : Boolean;
    TotHtLigne,TotTaxeLigne,TotTTCLigne,TotGlobal : Double;
    tmp, Champ,Valeur : string;
    X, NbEch : integer;
begin
// Le calcul se fait uniquement sur les valeurs dev le verifTom => maj et conversion
// Récup des montants lignes
Tmp:=(Trim(ReadTokenSt(RecupTotaux)));
totHtLigne :=0;
TotTTCligne :=0;
While (Tmp <> '') do
    begin
    X:=pos(':',Tmp);
    if x<>0 then begin Champ:=copy(Tmp,1,X-1); Valeur:=Copy (Tmp,X+1,length(Tmp)-X); end
            else Champ := Tmp;
    // Recup des arguments ...
    if Champ = 'GP_TOTALHTDEV'  then TotHtLigne  := strToFloat(Valeur);
    if Champ = 'GP_TOTALTTCDEV' then TotTTCLigne := strToFloat(Valeur);
    Tmp:=(Trim(ReadTokenSt(RecupTotaux)));
    end;
TotTaxeLigne := TotTTCLigne - TotHTLigne;

TobAffNew.PutValue('AFF_TOTALHTDEV',TotHtLigne);
TobAffNew.PutValue('AFF_TOTALTAXEDEV',TotTaxeLigne);
TobAffNew.PutValue('AFF_TOTALTTCDEV',TotTTCLigne);

// Calcul Total HT Global de l'affaire
TotGlobal := CalculTotalAffaire(TobAffNew.GetValue('AFF_AFFAIRE'),TobAffNew.GetValue('AFF_GENERAUTO'),
         TobAffNew.GetValue('AFF_TYPEPREVU'), TotHTLigne, Nil, NbEch, bBaseEch ,TobAffNew.GetValue('AFF_DATEDEBGENER'));
// mcd 08/02/02 il faut passer la date pour que le factaff soit pris en compte       TobAffNew.GetValue('AFF_TYPEPREVU'), TotHTLigne, Nil, NbEch, bBaseEch ,iDate1900);

// ATTENTION : Ces tests sont identiques à ceux de la fonction MajTotAffaireParLesDEV ci-dessous
// PENSER A LES METTRE A JOUR EN CAS DE MODIF
// le Total affaire est parfois alimenté depuis les lignes
if ( TobAffNew.Getvalue('AFF_AFFCOMPLETE')<> '-')
   and  ( TobAffNew.Getvalue('AFF_CALCTOTHTGLO')='X')
   and    (TotGlobal<>0)     // pour ne pas ecraser si calcul depuis echénace(cas où pas de ligne
                             // ce dernier test n'est par contre pas utile dans l'autre fct
   then TobAffNew.PutValue('AFF_TOTALHTGLODEV',TotGlobal);
// Calcul du montant des écheances
(* Enleve LS
If (TobAffNew.GetValue('AFF_GENERAUTO')= 'CON')  then
   TobAffNew.PutValue ('AFF_MONTANTECHEDEV' ,TobAffNew.GetValue('AFF_TOTALHTDEV')*TobAffNew.GetValue('AFF_INTERVALGENER'));
*)
// FIN DE L''ATTENTION
end;

Procedure MajTotAffaireParLesDEV (TobAff : TOB; bTenueEuro:boolean);
BEGIN
TobAff.PutValue('AFF_TOTALHT',TobAff.GetValue('AFF_TOTALHTDEV'));
TobAff.PutValue('AFF_TOTALTAXE',TobAff.GetValue('AFF_TOTALTAXEDEV'));
TobAff.PutValue('AFF_TOTALTTC',TobAff.GetValue('AFF_TOTALTTCDEV'));
// ATTENTION : Ces tests sont identiques à ceux de la fonction RecalculTotAffaire ci-dessus
// PENSER A LES METTRE A JOUR EN CAS DE MODIF
// le Total affaire est parfois alimenté depuis les lignes
if ( TobAff.Getvalue('AFF_AFFCOMPLETE')<> '-')
   and  ( TobAff.Getvalue('AFF_CALCTOTHTGLO')='X')
   then begin
          TobAff.PutValue('AFF_TOTALHTGLO',TobAff.GetValue('AFF_TOTALHTGLODEV'));
        end;
// Calcul du montant des écheances
If (TobAff.GetValue('AFF_GENERAUTO')= 'CON')  then
        begin
          TobAff.PutValue ('AFF_MONTANTECHE' ,TobAff.GetValue('AFF_MONTANTECHEDEV'));
        end;
// FIN DE L''ATTENTION


end;


Procedure InitAppreciation (TobDet : TOB);
begin
tobDet.PutValue('AFA_APPRECIEE', '-') ;
tobDet.PutValue('AFA_ECHEFACT', '-') ;
tobDet.PutValue('AFA_TYPECHE', 'NOR') ;
TobDet.PutValue('AFA_BONIMALI',  0) ;
TobDet.PutValue('AFA_BONIMALIDEV',  0) ;
TobDet.PutValue('AFA_AFACTURER',  0) ;
TobDet.PutValue('AFA_AFACTURERDEV',  0) ;
TobDet.PutValue('AFA_NUMPIECE', '') ;
TobDet.PutValue('AFA_BLOCNOTE', '') ;
TobDet.PutValue('AFA_JUSTIFBONI', '') ;
TobDet.PutValue('AFA_ETATVISA', '') ;
TobDet.PutValue('AFA_VISEUR', '') ;
TobDet.PutValue('AFA_DATEVISA', idate1900) ;
   // 562 ajout zones manquantes ..
TobDet.PutValue('AFA_JUSTIFBONI2', '') ;
TobDet.PutValue('AFA_LIBELLE1', '') ;
TobDet.PutValue('AFA_LIBELLE2', '') ;
TobDet.PutValue('AFA_LIBELLE3', '') ;
TobDet.PutValue('AFA_BONIMALIQTE',  0) ;
TobDet.PutValue('AFA_BM2PRQTE',  0) ;
TobDet.PutValue('AFA_AFACTURERQTE',  0) ;
TobDet.PutValue('AFA_ACPTEQTE',  0) ;
TobDet.PutValue('AFA_BM1FO',  0) ;
TobDet.PutValue('AFA_BM1FODEV',  0) ;
TobDet.PutValue('AFA_BM1FR',  0) ;
TobDet.PutValue('AFA_BM1FRDEV',  0) ;
TobDet.PutValue('AFA_BM2PR',  0) ;
TobDet.PutValue('AFA_BM2PRDEV',  0) ;
TobDet.PutValue('AFA_BM2FO',  0) ;
TobDet.PutValue('AFA_BM2FODEV',  0) ;
TobDet.PutValue('AFA_BM2FR',  0) ;
TobDet.PutValue('AFA_BM2FRDEV',  0) ;
TobDet.PutValue('AFA_AFACTOT',  0) ;
TobDet.PutValue('AFA_AFACTOTDEV',  0) ;
TobDet.PutValue('AFA_AFACTFR',  0) ;
TobDet.PutValue('AFA_AFACTFRDEV',  0) ;
TobDet.PutValue('AFA_AFACTFO',  0) ;
TobDet.PutValue('AFA_AFACTFODEV',  0) ;
TobDet.PutValue('AFA_ACPTEPR',  0) ;
TobDet.PutValue('AFA_ACPTEPRDEV',  0) ;
TobDet.PutValue('AFA_ACPTEFO',  0) ;
TobDet.PutValue('AFA_ACPTEFODEV',  0) ;
TobDet.PutValue('AFA_ACPTEFR',  0) ;
TobDet.PutValue('AFA_ACPTEFRDEV',  0) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 05/09/2002
Modifié le ... :   /  /
Description .. : Duplication des taches d'une affaire
Mots clefs ... :
*****************************************************************}
Function DuplicationTache(pStAffaireSource, pStAffaireDest : String) : Boolean;
var
  vSt         : String;
  vTob        : Tob;
  vQR         : TQuery;
  vStAff0     : String;
  vStAff1     : String;
  vStAff2     : String;
  vStAff3     : String;
  vStAvenant  : String;
  vStTiers    : String;
  i           : Integer;
  
begin
  result := true;
    //** duplication, il faut tout prendre ...
  vSt := 'SELECT * FROM TACHE WHERE ATA_AFFAIRE = "' + pStAffaireSource + '"';
  vQr := nil;
  Try
    vQR := OpenSql(vSt,True,-1,'',true);
    if Not vQR.Eof then
      begin
        vTob := Tob.create('MesTaches', nil, -1);
        Try                              
          vTob.LoadDetailDB('TACHE','','',vQR,False,True);
          // changement d'affaire
          for i := 0 to vTob.detail.count - 1 do
            begin

		      {$IFDEF BTP}
              BTPCodeAffaireDecoupe(pStAffaireDest, vStAff0, vStAff1, vStAff2, vStAff3, vStAvenant, taModif, false);
		      {$ELSE}
              CodeAffaireDecoupe(pStAffaireDest, vStAff0, vStAff1, vStAff2, vStAff3, vStAvenant, taModif, false);
		      {$ENDIF}

              TeststCleAffaire(pStAffaireDest,vStAff0,vStAff1,vStAff2,vStAff3,vStAvenant,vStTiers, False, False, False, False);

              vTob.detail[i].putValue('ATA_AFFAIRE', pStAffaireDest);
              vTob.detail[i].putValue('ATA_AFFAIRE0', vStAff0);
              vTob.detail[i].putValue('ATA_AFFAIRE1', vStAff1);
              vTob.detail[i].putValue('ATA_AFFAIRE2', vStAff2);
              vTob.detail[i].putValue('ATA_AFFAIRE3', vStAff3);
              vTob.detail[i].putValue('ATA_AVENANT', vStAvenant);
              vTob.detail[i].putValue('ATA_TIERS', vStTiers);
              vTob.detail[i].putValue('ATA_NUMEROTACHE', intToStr(i + 1));
            end;
          vTob.InsertDB(nil, False);
        Finally
          vTob.free;                
        End;
      end;
  finally
    ferme(vQr);
  end;
end;

end.
