{Source de la tof CPMULCFONB
--------------------------------------------------------------------------------------
  Version     |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
07.01.001.001  03/07/06   JP   création de l'unité

--------------------------------------------------------------------------------------}
unit CPMULCFONB;

interface

uses
    StdCtrls, Windows, Controls, Classes, uTob,
{$IFDEF EAGLCLIENT}
    eMul, MaineAGL,
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  db, HDB, mul, FE_Main,
{$ENDIF}
    Forms, SysUtils, ComCtrls, HCtrls, HEnt1, HMsgBox, HTB97, Ent1,
    UTOF, LookUp, uLibCFONB,paramsoc;

type
  TOF_CPMULCFONB = class(TOF)
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S : string); override;
    procedure OnClose; override;
  private
    ObjCFONB      : TObjCFONB;
    Libelle,Adresse1,Adresse2,CodePostal,CodeNif,NumeroTVA,Ville : string;
    GetInfoParamSoc : boolean;
    procedure CreateObjCFONB;
    {Mise à jour de la DBListe en focntion du type d'export}
    procedure SetLaListe    (TypCFONB : string);
    {Récupère la valeur d'un champ de la grille}
    function  GetValeur     (NomChamp : string) : Variant;
    {Création d'une tob à partir d'un nom de liste}
    function  NewTobFromDBL (NomListe : string) : TOB;
    {Rajout des infos de la table PARAMSOC à une TOB fille}
    procedure GetInfosParamSoc (var TOBFille : TOB);
  public
    {$IFDEF EAGLCLIENT}
    FListe : THGrid;
    {$ELSE}
    FListe : THDBGrid;
    {$ENDIF}

    procedure CBTypeOnChange(Sender : TObject);  
    procedure PaysOnChange  (Sender : TObject);
    procedure BValiderClick (Sender : TObject);
    procedure GrilleDblClick(Sender : TObject);
    procedure SlctAllClick  (Sender : TObject);
    procedure BParamExpClick(Sender : TObject);
    procedure AuxiElipsisClick(Sender : TObject);
  end;

procedure CP_LanceFicheMulCFONB(Arg : string);


implementation

uses
  commun, CPSAISIEPIECE_TOF, CPVISUPIECES_TOF, SaisBor, LP_Param, LP_Base, AglInit, UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }
{---------------------------------------------------------------------------------------}
procedure CP_LanceFicheMulCFONB(Arg : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPMULCFONB', '', '', Arg);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCFONB.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin          
  inherited;
  GetInfoParamSoc := false;
  Ecran.HelpContext := 150;

  { BVE 18.04.07
    Si pas de CISXPGI defini on supprime l'option d'export via CISX
  }
  {$IFDEF CISXPGI}
  THValComboBox(GetControl('CBTYPE')).Plus := ' AND CO_CODE IN ("' + cfonb_Transfert + '", "' + cfonb_CISX + '")';
  {$ELSE}
  {Dans un premier temps, on ne gère que les virements internationaux
   JP 06/07/07 : Il faut maintenir ce code}
  if S = '' then begin
    SetControlText('CBTYPE', cfonb_Transfert);
    SetControlEnabled('LBTYPE', False);
    SetControlEnabled('CBTYPE', False);
  end;

  Ecran.Caption := TraduireMemoire('Export paramétrable');
  UpdateCaption(Ecran);

  THValComboBox(GetControl('CBTYPE')).Plus := ' AND CO_CODE <> "' + cfonb_CISX + '" ';
  {$ENDIF}

  {JP 29/01/07 : gestion du partage de BanqueCP : Uilisation d'une fonction générique}
  SetPlusBanqueCp(GetControl('CBBANQUE'));

  SetControlVisible('BSELECTALL', True);

  TToolbarButton97(GetControl('BOUVRIR'     )).OnClick := BValiderClick;
  TToolbarButton97(GetControl('BSELECTALL'  )).OnClick := SlctAllClick;
  TToolbarButton97(GetControl('BPARAMEXPORT')).OnClick := BParamExpClick;
  THValComboBox(GetControl('CBTYPE')).OnChange := CBTypeOnChange;
  CBTypeOnChange(GetControl('CBTYPE'));           
  THValComboBox(GetControl('PAYS')).OnChange := PaysOnChange;
  {$IFDEF EAGLCLIENT}
  FListe := THGrid(TFMul(Ecran).FListe);
  {$ELSE}
  FListe := THDBGrid(TFMul(Ecran).FListe);
  {$ENDIF}

  FListe.OnDblClick := GrilleDblClick;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('T_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('T_AUXILIAIRE_', true)).OnElipsisClick:=AuxiElipsisClick;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCFONB.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ObjCFONB) then FreeAndNil(ObjCFONB);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCFONB.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCFONB.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCFONB.BParamExpClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Etat : string;
begin
  {Pour forcer la mise à jour de la sélection du modèle d'export}
  NextPrevControl(Ecran);

  if not ExJaiLeDroitConcept(ccParamEtat, True) then Exit;

  Etat := GetControlText('MODELEEXPORT');

  Param_LPTexte(nil, 'F', GetControlText('CBTYPE'), Etat, True);
  LibereSauvGardeLP;

  {Rechargement de la Combo}
  THValComboBox(GetControl('MODELEEXPORT')).ReLoad;
  if Etat <> '' then SetControlText('MODELEEXPORT', Etat);
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULCFONB.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCFONB.GrilleDblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  lMulQ : TQuery;
begin
  {$IFDEF EAGLCLIENT}
  lMulQ := TFMul(Ecran).Q.TQ;
  lMulQ.Seek(FListe.Row - 1);
  {$ELSE}
  lMulQ := TFMul(Ecran).Q;
  {$ENDIF}
  
  if ((GetField('J_MODESAISIE') <> '-') and (GetField('J_MODESAISIE') <> '')) then
    LanceSaisieFolio(lMulQ, taConsult)
  else
    TrouveEtLanceSaisieParam(lMulQ, taConsult, GetField('E_QUALIFPIECE'), False, '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCFONB.BValiderClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  {Aucune sélection, on sort}
  if (FListe.NbSelected = 0)
  {$IFNDEF EAGLCLIENT}
  and not FListe.AllSelected
  {$ENDIF}
  then begin
    HShowMessage('0;' + Ecran.Caption + '; Aucun élément n''est sélectionné.;W;O;O;O;', '', '');
    Exit;
  end;

  if GetControlText('CBBANQUE') = '' then begin
    HShowMessage('0;' + Ecran.Caption + '; Veuillez choisir une banque.;W;O;O;O;', '', '');
    Exit;
  end;

  {Pas besoin de modele pour l'export vers CISX}
  if (GetControlText('CBTYPE') <> cfonb_CISX) and (GetControlText('MODELEEXPORT') = '') then begin
    HShowMessage('0;' + Ecran.Caption + '; Veuillez choisir un modèle d''export.;W;O;O;O;', '', '');
    Exit;
  end
  else if (GetControlText('CBTYPE') = cfonb_CISX) and (GetControlText('LISTESCRIPT') = '') then begin
    if THValComboBox(GetControl('LISTESCRIPT')).Values.Count = 0 then
       HShowMessage('0;' + Ecran.Caption + '; Pas de type d''export défini. Export impossible.;W;O;O;O;', '', '')
    else
       HShowMessage('0;' + Ecran.Caption + '; Veuillez choisir un type d''export.;W;O;O;O;', '', '');
    Exit;
  end;

  if Assigned(ObjCFONB) then FreeAndNil(ObjCFONB);
  CreateObjCFONB;

  if ObjCFONB.TypeExport = '' then begin
    HShowMessage('1;' + Ecran.Caption + '; Vous devez choisir un type d''export.;W;O;O;O;', '', '');
    SetFocusControl('CBTYPE');
    Exit;
  end;

  TFMul(Ecran).Q.First;
  if FListe.AllSelected then
    while not TFMul(Ecran).Q.EOF do begin
      {Récupération du paramétrage de la grille et de valeurs de la ligne courrante}
      NewTobFromDBL(TFMul(Ecran).FNomFiltre);
      TFMul(Ecran).Q.Next;
    end

  else

    for n := 0 to FListe.nbSelected - 1 do begin
      FListe.GotoLeBookmark(n);
      {$IFDEF EAGLCLIENT}
      TFMul(Ecran).Q.TQ.Seek(FListe.Row - 1);
      {$ENDIF}
      {Récupération du paramétrage de la grille et de valeurs de la ligne courrante}
      NewTobFromDBL(TFMul(Ecran).FNomFiltre);
    end;

  { Traitement particulier si export vers CISX }
  if ObjCFONB.TypeExport = cfonb_CISX then
  begin
     ObjCFONB.Pays := GetControlText('PAYS');
     ObjCFONB.Script := GetControlText('LISTESCRIPT');
     ObjCFONB.Banque := GetControlText('CBBANQUE');
     ObjCFONB.LanceTraitement;
  end
  else
  begin
    {Récupération des informations generiques du fichier CFONB}
    ObjCFONB.GetInfosSaisies;
    {Complétion de la tob de traitement}
    ObjCFONB.LanceTraitement;
    {Lancement éventuel d'un bordereau}
    ObjCFONB.LanceBordereau;
    {Affichage des événtuels rejets}
    if ObjCFONB.TobEcrRejetees.Detail.Count > 0 then begin
       if HShowMessage('0;' + Ecran.Caption + '; Voulez-vous voir la liste des écritures rejetées ?;I;YN;Y;N;', '', '') = mrYes then begin
         TheTob := ObjCFONB.TobEcrRejetees;
         CPLanceFiche_VisuPieces('ACTION=TACONSULT;TITRE=LISTE DES ÉCRITURES REJETÉES;CHAMPS=' +
          'SYSDOSSIER,E_ETABLISSEMENT,E_JOURNAL,E_NUMEROPIECE,E_DATECOMPTABLE,E_REFINTERNE,ERREUR;');
       end;
    end;
  end;

  {Raffraîchissement de la liste}
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCFONB.PaysOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
Var
  SQL : string;
  Combo : THValComboBox;
begin
  // Ne pouvant pas passer par une tablette on gere le chargement du THValComboBox.
  SQL := 'SELECT CIS_CLE,CIC_LIBELLE ' +
         'FROM CPGZIMPREQ, CPGZIMPCORRESP ' +
         'WHERE CIS_CLE<>"" AND ' +
         'CIS_CLE = CIC_IDENTIFIANT AND ' +
         'CIS_NATURE<> "X" AND ' +
         'CIS_TABLE3="' + GetControlText('PAYS') + '" AND ' +
         'CIS_CLEPAR="EXPORT" ' +
         'ORDER BY CIC_LIBELLE';
  Combo := THValComboBox(GetControl('LISTESCRIPT'));
  ChargeTHValComboBox(SQL,Combo);
  SetFocusControl('LISTESCRIPT');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCFONB.CBTypeOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Typ : string;
begin
  Typ := GetControlText('CBTYPE');
  {Changement de DBLISTE en fonction du type d'export}
  SetLaListe(Typ);
  { Selon le type d'export on affiche ou non certaines zones }
  if (Typ = cfonb_CISX) then
  begin
    SetControlVisible('MODELEXEPORT',false);
    SetControlVisible('LBLMODELEEXPORT',false);
    SetControlVisible('CBOMODELEBOR',false);
    SetControlVisible('LBLBORDEREAU',false);
    SetControlVisible('BPARAMEXPORT',false);
    SetControlVisible('PAYS',true);
    SetControlVisible('LBLPAYS',true);
    SetControlVisible('LISTESCRIPT',true);
    SetControlVisible('LBLLISTESCRIPT',true);
    THValComboBox(GetControl('PAYS')).Itemindex := THValComboBox(GetControl('PAYS')).Values.IndexOf(GetParamSocSecur('SO_PAYS',''));
    PaysOnChange(nil);
  end
  else
  begin 
    SetControlVisible('MODELEXEPORT',true);
    SetControlVisible('LBLMODELEEXPORT',true);
    SetControlVisible('CBOMODELEBOR',true);
    SetControlVisible('LBLBORDEREAU',true); 
    SetControlVisible('BPARAMEXPORT',true);
    SetControlVisible('PAYS',false);
    SetControlVisible('LBLPAYS',false);
    SetControlVisible('LISTESCRIPT',false);
    SetControlVisible('LBLLISTESCRIPT',false);
    {On filtre la tablette des modèles d'export}
    THValComboBox(GetControl('MODELEEXPORT')).Plus := Typ;
    THValComboBox(GetControl('MODELEEXPORT')).Refresh;
    THValComboBox(GetControl('MODELEEXPORT')).ItemIndex := 0;
  end;
end;

{Mise à jour de la DBListe en focntion du type d'export
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCFONB.SetLaListe(TypCFONB: string);
{---------------------------------------------------------------------------------------}
begin
       if TypCFONB = cfonb_Transfert   then TFMul(Ecran).SetDBListe('CPLISTETRACFONB')
       else if TypCFONB = cfonb_CISX   then TFMul(Ecran).SetDBListe('CPLISTECISXCFONB')
  {
  else if TypCFONB = cfonb_Virement    then TFMul(Ecran).SetDBListe('CPLISTEVIRCFONB')
  else if TypCFONB = cfonb_Prelevement then TFMul(Ecran).SetDBListe('CPLISTEPRECFONB')
  ...}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCFONB.SlctAllClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Fiche : TFMul;
begin
  Fiche := TFMul(Ecran);
  {$IFDEF EAGLCLIENT}
  if not Fiche.FListe.AllSelected then begin
    if not Fiche.FetchLesTous then Exit;
  end;
  {$ENDIF}
  Fiche.bSelectAllClick(nil);
end;

{Récupère la valeur d'un champ de la grille
{---------------------------------------------------------------------------------------}
function TOF_CPMULCFONB.GetValeur(NomChamp : string) : Variant;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TFMul(Ecran).Q.FindField(NomChamp)) then
    Result := TFMul(Ecran).Q.FindField(NomChamp).AsVariant
  else
    Result := '';
end;

{Création d'une tob à partir d'un nom de liste
{---------------------------------------------------------------------------------------}
function TOF_CPMULCFONB.NewTobFromDBL(NomListe: string): TOB;
{---------------------------------------------------------------------------------------}
var
  lStSource    : string;
  lStLiens     : string;
  lStTri       : string;
  lStChamps    : string;
  lStTitres    : hString;
  lStLargeurs  : string;
  lStJustifs   : string;
  lStParams    : string;
  lStLibelle   : hString;
  lStNumCols   : hString;
  lStPerso     : string;
  lBoOkTri     : Boolean;
  lBoOkNumCol  : Boolean;
  NomChamp     : string;
begin
  Result := TOB.Create(NomListe, ObjCFONB.TobTraitement, -1);

  ChargeHListe(NomListe, lStSource, lStLiens, lStTri, lStChamps, lStTitres, lStLargeurs,
               lStJustifs, lStParams, lStLibelle, lStNumCols, lStPerso, lBoOkTri, lBoOkNumCol);

  NomChamp := ReadTokenSt(lStChamps);
  while not (Trim(NomChamp) = '') do begin
    Result.AddChampSupValeur(NomChamp, GetValeur(NomChamp));
    NomChamp := ReadTokenSt(lStChamps);
  end;
  { Cas Spécifique envoi vers CISX }
  if ObjCFONB.TypeExport = cfonb_CISX then
  begin
       { Rajout des infos de la table ParamSoc }
       GetInfosParamSoc(Result);
  end
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 26/12/2006
Modifié le ... :   /  /    
Description .. : Rajout des infos de la table ParamSoc à une TOB fille
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULCFONB.GetInfosParamSoc (var TOBFille : TOB);
var
   SQL : String;
   Q : TQuery;
begin
    if not(GetInfoParamSoc) then
    begin
      { Récupération des valeurs }
         SQL := 'SELECT SOC_NOM, SOC_DATA FROM PARAMSOC WHERE '+
                'SOC_NOM = "SO_LIBELLE" OR '+
                'SOC_NOM = "SO_ADRESSE1" OR '+
                'SOC_NOM = "SO_ADRESSE2" OR '+
                'SOC_NOM = "SO_CODEPOSTAL" OR '+
                'SOC_NOM = "SO_NIF" OR '+
                'SOC_NOM = "SO_RVA" OR '+
                'SOC_NOM = "SO_VILLE"';

        Q :=  OpenSQL(SQL,True);
        try
          while not(Q.Eof) do
          begin
             if Q.FindField('SOC_NOM').AsString = 'SO_LIBELLE' then
                Libelle := Q.FindField('SOC_DATA').AsString
             else if Q.FindField('SOC_NOM').AsString = 'SO_ADRESSE1' then
                Adresse1 := Q.FindField('SOC_DATA').AsString
             else if Q.FindField('SOC_NOM').AsString = 'SO_ADRESSE2' then
                Adresse2 := Q.FindField('SOC_DATA').AsString
             else if Q.FindField('SOC_NOM').AsString = 'SO_CODEPOSTAL' then
                CodePostal := Q.FindField('SOC_DATA').AsString
             else if Q.FindField('SOC_NOM').AsString = 'SO_NIF' then
                CodeNif := Q.FindField('SOC_DATA').AsString  
             else if Q.FindField('SOC_NOM').AsString = 'SO_RVA' then
                NumeroTVA := Q.FindField('SOC_DATA').AsString 
             else if Q.FindField('SOC_NOM').AsString = 'SO_VILLE' then
                Ville := Q.FindField('SOC_DATA').AsString;
             Q.Next;
          end
        finally
          Ferme(Q);
        end;
        GetInfoParamSoc:=true;
    end;
    TOBFille.AddChampSupValeur('SO_LIBELLE',Libelle);
    TOBFille.AddChampSupValeur('SO_ADRESSE1',Adresse1);
    TOBFille.AddChampSupValeur('SO_ADRESSE2',Adresse2);
    TOBFille.AddChampSupValeur('SO_CODEPOSTAL',CodePostal);
    TOBFille.AddChampSupValeur('SO_NIF',CodeNif);   
    TOBFille.AddChampSupValeur('SO_TVA',NumeroTVA);
    TOBFille.AddChampSupValeur('SO_VILLE',Ville);
end;
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCFONB.CreateObjCFONB;
{---------------------------------------------------------------------------------------}
var
  Typ : string;
begin
  Typ := GetControlText('CBTYPE');
       if Typ = cfonb_Transfert   then ObjCFONB := TObjTransfert  .Create('oooo', nil, -1)
  else if Typ = cfonb_Virement    then ObjCFONB := TObjVirement   .Create('oooo', nil, -1)
{$IFDEF CISXPGI}
  else if Typ = cfonb_CISX        then ObjCFONB := TObjCISX       .Create('oooo', nil, -1)
{$ENDIF}  
  else if Typ = cfonb_Prelevement then ObjCFONB := TObjPrelevement.Create('oooo', nil, -1);

  ObjCFONB.TypeExport := Typ;
  ObjCFONB.SetString('COMPTEBQ', GetControlText('CBBANQUE'));
  ObjCFONB.SetString('DOCUMENT', GetControlText('CBOMODELEBOR'));
  ObjCFONB.SetString('MODELEFICHIER', GetControlText('MODELEEXPORT'));
end;


initialization
  RegisterClasses([TOF_CPMULCFONB]);

end.
