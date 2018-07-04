unit Outils;

// CA - 18/05/99 -
// TGA - 06/10/05 - gestion dépréciation d'actif pour maj ope en cours
// BTY - 10/05 - En MAJ ope en cours, cas particulier des opérations MBA/MB2
// BTY - 01/06 FQ 17259 Nouveau top dépréciation dans IMMO
// BTY - 04/06 FQ 17516 Nouveau top changement de regroupement dans IMMO
// BTY - 04/06 FQ 17451 Nelles opérations en type énuméré
// BTY - 11/06 FQ 19036 MAJ date fin de contrat des immos pour pouvoir les filtrer sur cette date en édition
// MBO - 17/04/2007 - nvelle fonction SupRemplacement qui renvoie ok
//                    si l'on peut annuler l'opération de remplacement de composant
// BTY - 05/07 FQ 20256 Type dérogatoire -> i_typederoglia
// MBO - 05/07/2007 - modif f° SupRemplacement : modif du test sur i_operation
// BTY - 07/07  Publication des fonctions de gestion des exercices relatifs en multicritères
//              viennent de AMLISTE_TOF
// MBO - 19/07/07 FQ 21151 - correction message controle durée minimum en linéaire
// BTY 08/10/07 - FQ 21609 - Echéancier : NB échéances fausses + Erreur SQL clé en double avec NOMBREMOIS appelé par RecalculTranches->CalculNombreEcheance
// => j'ai remplacé NOMBREMOIS par IntervalleNbMois dans CalculNombreEcheance

interface

uses  SysUtils,
      HCtrls,
      HEnt1,
      Utob,
      Controls,  // pour les TControl des exos relatifs
      {$IFDEF VER150}
      Variants,
      {$ENDIF}
      {$IFDEF eAGLClient}
      {$ELSE}
      {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
      {$ENDIF eAGLClient}
      ImEnt,
      StdCtrls,   // Pour TEdit des exos relatifs
      {$IFDEF SERIE1}
      S1Util; // pour TexoDate des exos relatifs
      {$ELSE}
      {$IFDEF MODENT1}
      CPTypeCons,
      {$ENDIF MODENT1}
      Ent1;   // pour TexoDate des exos relatifs
      {$ENDIF}



type TypeOperation = (toNone,toCession,toCPartielle,toEclat,toMutation,toDuree,
                      toMethode,toAjoutDot,toSubDot,toReintegre,toAnnulation,
                      toRachat,toGenEcr,toChangePlan,toLeveeOption,toChanEtabl,
                      toChanLieu, toRegroupement,
                      toRevision,toExceptionnel,toModifBases,toDeprec) ;
                      // BTY 04/06 FQ 17451
                      // toChanLieu, toRegroupement) ;
// Gestion des ARD
Type TypeDotation = (toDotMin, toDotMax, toDotTheorique, toDotGlobale, toDotImmo);

function MontantToStr (Valeur : extended) : string;
function NouveauCodeImmo(Ajout:integer=0) : string ;
function ExisteCodeImmo(Code : string) : boolean ;
procedure OuvrirFichierIE(var FichierIE : TextFile;NomFichier : string;RecSize : word);
procedure EcrireLigneIE(var FichierIE : TextFile; Ligne : string);
procedure FermerFichierIE(var FichierIE : TextFile);
function TrouveNumeroOrdreLogSuivant(CodeImmo : string):integer;
function TrouveNumeroOrdreSerieLogSuivant:integer;
//function IsOpeEnCours (Code : string) : boolean;
function IsOpeEnCours (Q : TQuery; Code : string; bExerciceEnCours : boolean) : boolean;
function AffecteCommentaireOperation ( TypeOpe : string): string;

{$IFDEF eAGLClient}
{$ELSE}
// BTY 01/06 FQ 17259 Nouveau top de dépréciation dans IMMO
// BTY 04/06 FQ 17516 Nouveau top changement de regroupement
//procedure InitOpeEnCoursImmo(var Q : TQuery ; bMu,bEc,bCe,bCp,bLg,bEt,bLe,bMb,bGe: string); overload
//procedure InitOpeEnCoursImmo(var Q : TQuery ; bMu,bEc,bCe,bCp,bLg,bEt,bLe,bMb,bDp,bGe: string); overload
procedure InitOpeEnCoursImmo(var Q : TQuery ; bMu,bEc,bCe,bCp,bLg,bEt,bLe,bMb,bDp,bRg,bGe: string); overload
{$ENDIF}
//procedure InitOpeEnCoursImmo(T : TOB ; bMu,bEc,bCe,bCp,bLg,bEt,bLe,bMb,bGe: string); overload
//procedure InitOpeEnCoursImmo(T : TOB ; bMu,bEc,bCe,bCp,bLg,bEt,bLe,bMb,bDp,bGe: string); overload
procedure InitOpeEnCoursImmo(T : TOB ; bMu,bEc,bCe,bCp,bLg,bEt,bLe,bMb,bDp,bRg,bGe: string); overload
procedure MajOpeEnCoursImmo ( var QImmo : TQuery; Champ, TypeOpe, Coche : string);
function TraiteIntervaleDureeAmort(Methode : string;Orig, Valeur : integer;var NumMess : integer) : boolean;
function ImmoAmortie(MethodeEco,MethodeFisc,Nature : string) : boolean;
procedure CocheChampOperation (Q : TQuery; Code,Champ : string);
function MinDureeAmortissement ( stMethode : string; iDuree : integer; var stMsg : string ) : boolean;
function CalculNombreEcheance (DateDebut, DateFin : TDateTime; Periode : integer) : integer;  //YCP 25/08/05
procedure IntervalleNbMois (D1, D2 : TDateTime ; var Nb: Word);

procedure MAJDateFinContratImmos ;

function SupRemplacement( CodeOrigine : string): boolean;
function TypeDerogatoire( T:Tob ; Q:TQuery): string;  // FQ 20256


// Gestion des exercices relatifs présentés en THValComboBox d'un multicritères
{$IFNDEF eAGLServer}
procedure AMInitComboExercice (ComBo : THValComboBox );
procedure AMAppliquerExoRelatifToDates (CodeExoRelatif : string;
            TDateDebut, TDateFin : TControl ; bFiltre : boolean = False ) ;
{$ENDIF}


implementation
uses ImContra, ParamSoc;  // Pour MAJDateFinContratImmos

// -**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-*
// MAJ i_datefincb pour les crédit-bails et locations
// -**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-*
procedure MAJDateFinContratImmos ;
var  Q: TQuery;
     wContrat: TImContrat;
begin

  if (GetParamSocSecur ('SO_CALFINCONTRAT', iDate1900) = iDate1900) then
  begin
    Q:=OpenSql
    ('SELECT * FROM IMMO WHERE (I_NATUREIMMO="CB" OR I_NATUREIMMO="LOC") AND I_QUALIFIMMO <> "REG" ORDER BY I_IMMO', False);
    while (not Q.Eof) do
    begin
      try
      if (Q.FindField ('I_DATEFINCB').AsDateTime = iDate1900)
      or VarIsNull(Q.FindField('I_DATEFINCB').AsVariant) then
         begin
         wContrat := TImContrat.Create;
         wContrat.Charge(Q);
         wContrat.ChargeTableEcheance;
         wContrat.CalculEcheances;
         Q.Edit;
         Q.FindField ('I_DATEFINCB').AsDateTime := wContrat.GetDateFinContrat;
         Q.Post;
         FreeAndNil(wContrat);
         end;
      except
      end ;
      Q.Next ;
    end ;
    Ferme(Q) ;
    SetParamSoc ('SO_CALFINCONTRAT', Now);
  end;
end;


// -**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-*
// Formatage d'un montant
// -**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-*
function MontantToStr (Valeur : extended) : string;
begin
  result := StrfMontant ( Valeur,20, V_PGI.OkDecV , '', false)
end;

// -**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-*
// Récupération d'un code Immo automatique
// -**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-**-*
function NouveauCodeImmo(Ajout:integer=0) : string ;
var Q : TQuery ; iCode: int64 ;
begin
  iCode:=1 ;
  Q:=OpenSql('SELECT I_IMMO FROM IMMO ORDER BY I_IMMO DESC',true) ;
  while (not Q.Eof) and (iCode=1) do
  begin
    try
      iCode:=StrToInt64(Q.Fields[0].AsString)+1+Ajout ;
    except
      icode:=1 ;
    end ;
    Q.Next ;
  end ;
  Ferme(Q) ;
  result:=Format('%.10d',[iCode]) ;
end ;

function ExisteCodeImmo(Code : string) : boolean ;
var Query : TQuery ;
begin
Result:=FALSE ;
Query:=OpenSQL('SELECT I_IMMO FROM IMMO WHERE I_CHANGECODE="'+Code+'" OR I_IMMO="'+Code+'"', FALSE) ;
if not Query.EOF then Result:=TRUE ;
end ;

procedure OuvrirFichierIE(var FichierIE : TextFile;NomFichier : string;RecSize : word);
begin
  AssignFile(FichierIE,NomFichier);
  Rewrite(FichierIE);
end;

procedure EcrireLigneIE(var FichierIE : TextFile; Ligne : string);
begin
  Writeln(FichierIE,Ligne);
	Flush(FichierIE);
end;

procedure FermerFichierIE(var FichierIE : TextFile);
begin
	CloseFile(FichierIE);
end;

function TrouveNumeroOrdreLogSuivant(CodeImmo : string):integer;
var QueryC : TQuery ;
begin
  QueryC:=OpenSQL('SELECT MAX(IL_ORDRE) FROM IMMOLOG WHERE IL_IMMO="'+CodeImmo+'"', TRUE) ;
  if QueryC.EOF then result:=1
  else result:=QueryC.Fields[0].AsInteger+1 ;
  Ferme(QueryC) ;
end;

function TrouveNumeroOrdreSerieLogSuivant:integer;
var QueryC : TQuery ;
begin
  QueryC:=OpenSQL('SELECT MAX(IL_ORDRESERIE) FROM IMMOLOG', TRUE) ;
  if QueryC.EOF then result:=1
  else result:=QueryC.Fields[0].AsInteger+1 ;
  Ferme(QueryC) ;
end;

{function IsOpeEnCours (Code : string) : boolean;
var Q : TQuery;
begin
  Q:=OpenSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+Code+'"',True);
  if not Q.EOF then Result := True
  else Result := False;
  Ferme (Q);
end;}

{-------------------------------------------------------------------------------
   Function IsOpeEnCours ( Q : TQuery ; - Enreg IMMO, nil si passage par le code
                           Code : string; - Code de l'immo à vérifier
                           bExerciceEnCours : boolean - Exercice en cours uniqmt)
 Cette fonction permet de vérifier si des opérations sont en cours sur l'immo.
 considérée.
-------------------------------------------------------------------------------}
function IsOpeEnCours (Q : TQuery; Code : string; bExerciceEnCours : boolean) : boolean;
var select, where : string;
begin
  if (Q = nil) or (not bExerciceEnCours) then
  begin
    select := 'SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+Code+'"';
    if bExerciceEnCours then
      where := ' AND IL_DATEOP <="'+USDateTime(VHImmo^.Encours.Fin)+
               '" AND IL_DATEOP >="'+USDateTime(VHImmo^.Encours.Deb)+'"'
    else where := '';
    Q:=OpenSQL (select+where+ ' AND IL_TYPEOP<>"ACQ" AND IL_TYPEOP<>"CLO"',True);
    if not Q.EOF then Result := True
    else Result := False;
    Ferme (Q);
  end
  else  Result := (Q.FindField ('I_OPERATION').AsString = 'X');
end;

procedure CocheChampOperation (Q : TQuery; Code,Champ : string);
begin
  if Q <> nil then
  begin
    Q.FindField (Champ).AsString := 'X';
    Q.FindField ('I_OPERATION').AsString := 'X';
  end
  else
    ExecuteSQL('UPDATE IMMO SET '+Champ+'="X",I_OPERATION="X" WHERE I_IMMO="'+Code+'"') ;
end;

function AffecteCommentaireOperation (TypeOpe : string): string;
begin
if TypeOpe='CHP' then result := 'Plan révisé par '+V_PGI.UserName+ ' le '+ DatetoStr(Date)
else result := 'Effectué par '+V_PGI.UserName+ ' le '+ DatetoStr(Date);
end;

{$IFDEF eAGLClient}
{$ELSE}
//procedure InitOpeEnCoursImmo(var Q : TQuery ; bMu,bEc,bCe,bCp,bLg,bEt,bLe,bMb,bGe: string);
//procedure InitOpeEnCoursImmo(var Q : TQuery ; bMu,bEc,bCe,bCp,bLg,bEt,bLe,bMb,bDp,bGe: string);
procedure InitOpeEnCoursImmo(var Q : TQuery ; bMu,bEc,bCe,bCp,bLg,bEt,bLe,bMb,bDp,bRg,bGe: string);begin
  if bMu<>'' then Q.FindField('I_OPEMUTATION').AsString:=bMu;
  if bEc<>'' then Q.FindField('I_OPEECLATEMENT').AsString:=bEc;
  if bCe<>'' then Q.FindField('I_OPECESSION').AsString:=bCe;
  if bCp<>'' then Q.FindField('I_OPECHANGEPLAN').AsString:=bCp;
  if bLg<>'' then Q.FindField('I_OPELIEUGEO').AsString:=bLg;
  if bEt<>'' then Q.FindField('I_OPEETABLISSEMENT').AsString:=bEt;
  if bLe<>'' then Q.FindField('I_OPELEVEEOPTION').AsString:=bLe;
  if bMb<>'' then Q.FindField('I_OPEMODIFBASES').AsString:=bMb;
  // BTY 01/06 FQ 17259 Nouveau top dépréciation
  if bDp<>'' then Q.FindField('I_OPEDEPREC').AsString:=bDp;
  // BTY 04/06 FQ 17516 Nouveau top changement de regroupement
  if bRg<>'' then Q.FindField('I_OPEREG').AsString:=bRg;
  if bGe<>'' then Q.FindField('I_OPERATION').AsString:=bGe;
end;
{$ENDIF}

//procedure InitOpeEnCoursImmo(T : TOB ; bMu,bEc,bCe,bCp,bLg,bEt,bLe,bMb,bGe: string);
//procedure InitOpeEnCoursImmo(T : TOB ; bMu,bEc,bCe,bCp,bLg,bEt,bLe,bMb,bDp,bGe: string);
procedure InitOpeEnCoursImmo(T : TOB ; bMu,bEc,bCe,bCp,bLg,bEt,bLe,bMb,bDp,bRg,bGe: string);
begin
  if bMu<>'' then T.PutValue('I_OPEMUTATION',bMu);
  if bEc<>'' then T.PutValue('I_OPEECLATEMENT',bEc);
  if bCe<>'' then T.PutValue('I_OPECESSION',bCe);
  if bCp<>'' then T.PutValue('I_OPECHANGEPLAN',bCp);
  if bLg<>'' then T.PutValue('I_OPELIEUGEO',bLg);
  if bEt<>'' then T.PutValue('I_OPEETABLISSEMENT',bEt);
  if bLe<>'' then T.PutValue('I_OPELEVEEOPTION',bLe);
  if bMb<>'' then T.PutValue('I_OPEMODIFBASES',bMb);
  // BTY 01/06 FQ 17259 Nouveau top dépréciation
  if bDp<>'' then T.PutValue('I_OPEDEPREC',bDp);
  // BTY 04/06 FQ 17516 Nouveau top changement de regroupement
  if bRg<>'' then T.PutValue('I_OPEREG',bRg);
  if bGe<>'' then T.PutValue('I_OPERATION',bGe);
end;

procedure MajOpeEnCoursImmo ( var QImmo : TQuery; Champ, TypeOpe, Coche : string);
var Q: TQuery;
begin
{  Q := OpenSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+QImmo.FindField('I_IMMO').AsString+
                '" AND (IL_TYPEOP="'+TypeOpe+'")',True);
  if Q.RecordCount = 1 then QImmo.FindField(Champ).AsString := Coche;
  Ferme (Q);}
  Q := OpenSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+QImmo.FindField('I_IMMO').AsString+
  '" AND (IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'" AND IL_DATEOP<="'+USDateTime(VHImmo^.Encours.Fin)+
  '") AND (IL_TYPEOP<>"ACQ") AND (IL_TYPEOP<>"CLO")',True);
  if Q.RecordCount = 1 then
   begin
    QImmo.FindField(Champ).AsString := Coche;
    QImmo.FindField ('I_OPERATION').AsString := '-';
   end
  else
   begin
    // BTY 10/05 Cas particulier des opérations MBA/MB2
    if (TypeOpe = 'MBA') or (TypeOpe = 'MB2') then
      Q := OpenSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+QImmo.FindField('I_IMMO').AsString+
                   '" AND (IL_TYPEOP="MBA" OR IL_TYPEOP="MB2")'+
                   ' AND (IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+
                   '" AND IL_DATEOP<="'+USDateTime(VHImmo^.Encours.Fin)+
                   '") AND (IL_TYPEOP<>"ACQ") AND (IL_TYPEOP<>"CLO")', True)
    else
      Q := OpenSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+QImmo.FindField('I_IMMO').AsString+
                '" AND (IL_TYPEOP="'+TypeOpe+'")'+
        ' AND (IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'" AND IL_DATEOP<="'+USDateTime(VHImmo^.Encours.Fin)+
//        '")',True);
        '") AND (IL_TYPEOP<>"ACQ") AND (IL_TYPEOP<>"CLO")',True);

      if Q.RecordCount = 1 then QImmo.FindField(Champ).AsString := Coche;
  end;
  Ferme (Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 23/02/2004
Modifié le ... :   /  /    
Description .. : Contrôle que la durée d'amortissement est supérieure aux 
Suite ........ : limites inférieures autorisées. 
Suite ........ : Cette fonction renvoie le message d'erreur.
Mots clefs ... : 
*****************************************************************}
function MinDureeAmortissement ( stMethode : string; iDuree : integer; var stMsg : string ) : boolean;
const
  // FQ 21151 6 mbo - 19.07.07 MSG_LIN = 'La durée d''un amortissement linéaire doit être au moins égale à 12 mois';
  MSG_LIN = 'La durée d''un amortissement linéaire doit être au moins égale à 1 mois';
  MSG_DEG = 'La durée d''un amortissement dégressif doit être au moins égale à 36 mois';
//  MSG_VAR = 'La durée d''un amortissement variable doit être au moins égale à 12 mois';
var iDureeMin : integer;
begin
  stMSg := '';
  if (stMethode='DEG') then iDureeMin := 36
  else if (stMethode = 'LIN') then iDureeMin := 1
//  else if (stMethode = 'VAR') then iDureeMin := 12
  else iDureeMin := 0;
  if (iDuree < iDureeMin) then
  begin
    Result := False;
    if stMethode='DEG' then stMsg := MSG_DEG
//    else if stMethode='VAR' then stMsg := MSG_VAR
    else stMsg := MSG_LIN;
  end else result := true;
end;

function TraiteIntervaleDureeAmort(Methode : string;Orig, Valeur : integer;var NumMess : integer) : boolean;
var Min,Num1,Num2 : integer;
begin
// gère l'origine de l'appel pour un pb de numéro de message : imogen ou chanplan
  if Orig=1 then // appel depuis imogen
  begin Num1 := 33; Num2 := 34; end
  else begin Num1 := 7; Num2 := 8; end;
  if (Methode='DEG') then Min := 36
// CA<12  else if (Methode = 'LIN') then Min := 12
  else if (Methode = 'LIN') then Min := 1
  else Min := 0;
  if Valeur < Min then
  begin
    if (Methode='DEG') then NumMess := Num1
    else NumMess := Num2;
    result := false
  end
  else result := true;
end;

function ImmoAmortie(MethodeEco,MethodeFisc,Nature : string) : boolean;
begin
  result := (((MethodeEco<>'NAM') and (MethodeEco <> '')) or (MethodeFisc<>'')) and
             not ((Nature='LOC') or (Nature='FI'));
end;

function CalculNombreEcheance (DateDebut, DateFin : TDateTime; Periode : integer) : integer;  //YCP 25/08/05
//var nbEcheance : integer; laDate : TDateTime;
var NbMois : Word;
begin
{  nbEcheance := 0;
  laDate := DateDebut;
  while laDate<DateFin do
    begin
    Inc(nbEcheance,1);
    laDate := PlusMois (laDate,Periode);
    laDate := FINDEMOIS(LaDate);
    // si au cours de la boucle on passe sur un février,on n'a plus 30/numMois mais 28 ou 29/NumMois
  end;
  result:=nbEcheance+1;}
  //BTY 08/10/07 FQ 21609 NOMBREMOIS (DateDebut, DateFin, PremMois, PremAnnee, NbMois ) ;
  IntervalleNbMois (DateDebut, DateFin, NbMois );
  if (Periode > 1) then Result := (NbMois div Periode)
  else Result := NbMois;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 08/10/2007
Modifié le ... :   /  /
Description .. : Calcul de l'intervalle entre 2 dates, compté en nombre de mois
Suite ........ : Ex.  3 janv 07 au 2 fév 07   -> 1 mois
Suite ........ :      3 janv 07 au 3 fév 07   -> 2 mois
Mots clefs ... :
*****************************************************************}
procedure IntervalleNbMois (D1, D2 : TDateTime ; var Nb: Word);
var PremAnnee,PremMois,PremJour,DernAnnee,DernMois,DernJour : Word;
    Delta : integer;
begin
  if D1 > D2 then
  begin
    Nb := 0;
    exit;
  end;

  DecodeDate (D1,PremAnnee,PremMois,PremJour);
  DecodeDate (D2,DernAnnee,DernMois,DernJour);
  if DernJour < PremJour then  Delta := -1
  else Delta := 0;

  Nb := 12*(DernAnnee - PremAnnee) + (DernMois - PremMois) + 1 + Delta;

  {Contenu de la fonction CBP NOMBREMOIS :
  if D1 > D2 then
  begin
    NombreMois := 0;
    exit;
  end;
  DecodeDate (D1,PremAnnee,PremMois,PremJour);
  DecodeDate (D2,DernAnnee,DernMois,DernJour);
  NombreMois := 12 * (DernAnnee-PremAnnee) + (DernMois-PremMois+1); }
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 17/04/2007
Modifié le ... :   /  /
Description .. : fonction qui vérifie si l'on peut annuler une opération de
Suite ........ : remplacement de composant.
Suite ........ : On annule l'opération sur l'immo d'origine. On vérifie que le
Suite ........ : composant n'a pas été lui même remplacé.
Suite ........ : renvoie true si on peut supprimer
Suite ........ : renvoie false si on ne peut pas supprimer
Mots clefs ... :
*****************************************************************}
function SupRemplacement(CodeOrigine : string): boolean;
var
  Q : TQuery;
  CodeComposant : string;
  Condition : string;

begin

  result := true;
  Q:=OpenSQL('SELECT I_REMPLACEE, I_REMPLACE, I_OPERATION FROM IMMO WHERE I_IMMO="'+COdeOrigine+'"', FALSE) ;

  if Q.FindField('I_REMPLACEE').AsString <> '' then
  begin
    // c'est une immo remplacée - mbo 05.07.07 modif test sur i_operation (='-' modifié en <> "X")
    CodeComposant := Q.FindField('I_REMPLACEE').AsString;
    Condition := 'WHERE (I_IMMO ="' + CodeComposant;
    Condition := condition + '") AND (I_REMPLACE = "' + CodeOrigine;
    Condition := Condition + '") AND (I_OPERATION <> "X") ';
    Result := ExisteSQL ('SELECT * FROM IMMO ' + Condition);
  end;
  ferme(Q);
end;


function TypeDerogatoire( T:Tob ; Q:TQuery): string; // FQ 20256
begin
 Result := ' ';

 if (T = nil) then
 begin
   if Q.FindField ('I_METHODEFISC').AsString = 'DEG' then
      Result := 'DEG'
   else
      if (Q.FindField ('I_METHODEFISC').AsString <> '')
      and (Q.FindField ('I_DUREEECO').AsInteger > Q.FindField ('I_DUREEFISC').AsInteger) then
         Result := 'DUR'
      else
        if (Q.FindField ('I_METHODEFISC').AsString <> '') then
           Result := 'EXC';

 end else
 begin
   if T.GetValue ('I_METHODEFISC') = 'DEG' then
      Result := 'DEG'
   else
      if (T.GetValue ('I_METHODEFISC') <> '')
      and (T.GetValue ('I_DUREEECO') > T.GetValue ('I_DUREEFISC')) then
         Result := 'DUR'
      else
         if (T.GetValue ('I_METHODEFISC') <> '') then
           Result := 'EXC';
 end;
end;


{$IFNDEF eAGLServer}
{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 28/06/2006
Modifié le ... :   /  /
Description .. : FQ 18465  Pour l'exercice en cours VHImmo^.Encours,
Suite ........ : récup Position dans le tableau VHImmo^.Exercices
Suite ........ : Procédure locale NON PUBLIEE
Mots clefs ... :
*****************************************************************}
function AMPositionEncoursVHImmo : integer ;
var i: integer;
begin
  Result := -1;
  i := 1;
  while ((VHImmo^.Exercices[i].Code<>'') and (Result=-1)) do
  begin
    if VHImmo^.Exercices[i].Code=VHImmo^.Encours.Code then Result := i
    else Inc (i,1);
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 28/06/2006
Modifié le ... :   /  /
Description .. : FQ 18465 Calcul code relatif de l'exo passé en entrée
Suite ........ : Ce code est calculé selon sa chronologie par rapport à l'exo en cours
Suite ........ : =>  N+1, N+0, N-1 etc.
Suite ........ : Procédure locale NON PUBLIEE
Mots clefs ... :
*****************************************************************}
function AMCodeRelatifExo (PositionExo :integer; PositionEncours :integer) : string ;
begin
   Result := '';
   if PositionEncours = -1 then exit;
   if PositionExo < PositionEncours then Result := 'N-'+IntToStr(PositionEncours-PositionExo)
   else Result := 'N+'+IntToStr(PositionExo-PositionEncours);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 28/06/2006
Modifié le ... :   /  /
Description .. : FQ 18465 Récup code exercice connaissant son code relatif
Suite ........ : Procédure locale NON PUBLIEE
Mots clefs ... :
*****************************************************************}
function AMExoRelatifToExercice (CodeExoRelatif : string) : string ;
var
  iEncours, i : integer;
  stCodeRelatif : string;
begin
  Result := '';
  if CodeExoRelatif = '' then exit;
  if CodeExoRelatif[1]<>'N' then exit;

  // Récup rang exo en cours
  iEncours := AMPositionEncoursVHImmo;
  if iEnCours = - 1 then exit;

  // Récup code de l'exo relatif passé en entrée
  i := 1;
  while (VHImmo^.Exercices[i].Code<>'') do
  begin
    stCodeRelatif := AMCodeRelatifExo (i, iEncours);
    if stCodeRelatif = CodeExoRelatif then
    begin
      Result := VHImmo^.Exercices[i].Code;
      break;
    end;
    Inc(i,1);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 27/06/2006
Modifié le ... :
Description .. : ONGLET COMPLEMENTS
Suite ........ : FQ 18465 Libellés relatifs pour la liste des exos de la COMBO FExercice2
Suite ........ : -> N+1, N+0, N-1, N-2 etc. stockés dans COMBO.Value (non vu par l'utilisateur)
Suite ........ : -> Suivant:AAAA, En cours:AAAA, Précédent:AAAA, N-2:AAAA etc. stockés
Suite ........ : dans COMBO.Text (vu par l'utilisateur)
Mots clefs ... :
*****************************************************************}
procedure AMInitComboExercice ( ComBo : THValComboBox );

    // Formatage exercice : AAAA exo dans une seule année, AAAA/AAAA exo à cheval
    function FormaterLibelleExercice ( DateDebut, DateFin : TDateTime) : string;
    var yd, md, dd : Word;
        yf, mf, df : Word;
    begin
      DecodeDate ( DateDebut, yd, md, dd );
      DecodeDate ( DateFin, yf, mf, df);
      if yd = yf then Result := Format('%.04d',[yf])
      else Result := Format('%.04d/%.04d',[yd,yf]);
    end;

    // Libellé exo pour l'utilisateur (stocké dans les items de la COMBO)
    function LibelleExerciceRelatif ( stCode, stLibelle : string ) : string;
    begin
     if stCode = 'N-1' then Result :=  TraduireMemoire('Précédent')+' : '+stLibelle
     else if stCode = 'N+0' then Result :=  TraduireMemoire('En cours')+' : '+stLibelle
     else if stCode = 'N+1' then Result :=  TraduireMemoire('Suivant')+' : '+stLibelle
     else Result := stCode+' : '+stLibelle;
    end;

var
  iEncours, i : integer;
  stCode, stLibelle  : string;
  TExoLib, T : TOB;
begin
  if ComBo = nil then exit;
  ComBo.Items.Clear;
  ComBo.Values.Clear;
  // 1e ligne = <Tous>
  if ComBo.Vide then
  begin
    if ComBo.VideString = '' then
       ComBo.Items.Add(TraduireMemoire('<<Tous>>'))
    else ComBo.Items.Add(TraduireMemoire(ComBo.VideString));
    ComBo.Values.Add('');
  end;

  // Récup rang exo en cours dans le tableau VHImmo^.Exercices
  iEncours := AMPositionEncoursVHImmo;
  if iEnCours = - 1 then exit;

  // Chargement des exercices
  TExoLib := TOB.Create('', nil, -1);
  try
    TExoLib.LoadDetailDB('EXERCICE','','',nil,False);
    i:=1;
    while (VHImmo^.Exercices[i].Code<>'') do
    begin
      stCode := AMCodeRelatifExo (i, iEnCours);
      ComBo.Values.Add( stCode ); // code exo relatif dans le Value

      T := TExoLib.FindFirst(['EX_EXERCICE'],[VHImmo^.Exercices[i].Code],False);
      if T <> nil then
        stLibelle := FormaterLibelleExercice (VHImmo^.Exercices[i].Deb,VHImmo^.Exercices[i].Fin)
      else stLibelle := '';
      ComBo.Items.Add(LibelleExerciceRelatif ( stCode , stLibelle) ); // libellé exo relatif dans l'item

      Inc (i,1);
    end;
  finally
    TExoLib.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 27/06/2006
Modifié le ... :   /  /
Description .. : FQ 18465 Récup exercice réel à partir du code relatif N+1,N+0,N-1,N-2 etc.
Suite ........ : Ces dates réelles d'exo alimentent les TControl passés en entrée
Suite ........ : De plus, si on récupère un filtre utilisateur, on compare ces dates
Suite ........ : avec celles dans le filtre et on retient JJ MM si appartiennent à l'exo
Mots clefs ... :
*****************************************************************}
procedure AMAppliquerExoRelatifToDates (CodeExoRelatif : string;
          TDateDebut, TDateFin : TControl ; bFiltre : boolean = False ) ;
var
  stCode : string;
  DateDeb, DateFin, DateTmpDeb,DateTmpFin : TDateTime;
  Exo : TExoDate;
  yd, md, dd : word;
  ye, me, de : word;
begin

 // // Récup exercice réel à partir du code relatif N+1, N+0, N-1 etc.
 // // et alimentation des dates passées en entrée avec les dates de l'exercice
 // // ICI on n'a pas repris un filtre utilisateur
 if not bFiltre then
 begin

  // Récup code exo de l'exo relatif
  stCode := AMExoRelatifToExercice (CodeExoRelatif);
  if stCode = '' then exit;
  // Alimentation des TControl avec les dates exo
  // FQ 18465 07/06 ImExoDates au lieu de ExoToDates pour gérer le cas SERIE1
  IMExoToDates( stCode, TDateDebut, TDateFin);

 end else
 begin
 // // ICI récup filtre utilisateur contenant un exercice relatif :
 // // 1. récup exercice réel à partir du code relatif
 // // 2. comparaison des dates passées en entrée avec les dates de l'exercice recalculé

    stCode := AMExoRelatifToExercice ( CodeExoRelatif );
    if StCode = '' then Exit;

    if IsValidDate ( TEdit(TDateDebut).Text ) then
    begin
      DateDeb := StrToDate ( TEdit(TDateDebut).Text );
      if IsValidDate ( TEdit(TDateFin).Text ) then
      begin
        DateFin := StrToDate ( TEdit(TDateFin).Text );
        Exo.Code:= AMExoRelatifToExercice ( CodeExoRelatif );
        // Récup dates exo correspondant au code exo relatif en entrée
        // FQ 18465 07/06 ImRempliExoDate au lieu de RempliExoDate pour gérer le cas SERIE1
        IMRempliExoDate ( Exo );
        // Récup JJ MM AA du Tcontrol date début
        DecodeDate( DateDeb, yd, md, dd );
        DecodeDate( Exo.Deb, ye, me, de );
        // Remplacer JJ MM dans date début de l'exo par ceux du TControl en entrée
        DateTmpDeb := EncodeDate (ye, md, dd );
        // Récup JJ MM AA du TControl date fin
        DecodeDate( DateFin, yd, md, dd );
        DecodeDate( Exo.Fin, ye, me, de );
        // Remplacer JJ MM dans date fin de l'exo par ceux du TControl en entrée
        DateTmpFin := EncodeDate (ye, md, dd );
        if ((DateTmpDeb >=Exo.Deb) and (DateTmpDeb<=Exo.Fin)) and
              ((DateTmpFin >=Exo.Deb) and (DateTmpFin<=Exo.Fin)) then
        begin
          // JJ MM appartiennent à l'exo => les conserver
          TEdit(TDateDebut).Text := DateToStr( DateTmpDeb );
          TEdit(TDateFin).Text := DateToStr( DateTmpFin );
        end else
        begin
          // JJ MM n'appartiennent pas à l'exo => prendre les dates de l'exo
          TEdit(TDateDebut).Text := DateToStr( Exo.Deb );
          TEdit(TDateFin).Text := DateToStr( Exo.Fin );
        end;
      end;
    end; //ED1.text
 end;
end;
{$ENDIF}

end.

