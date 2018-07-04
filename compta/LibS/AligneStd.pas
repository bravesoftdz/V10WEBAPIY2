unit AligneStd;

interface

uses
      {$IFNDEF EAGLCLIENT}
            {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
      {$ENDIF}
      {$IFDEF VER150} Variants, {$ENDIF}
      Ent1,
      HMsgBox,
      HCtrls,
      SysUtils,
      uTOB,
      HEnt1,
      ParamSoc,
      ED_TOOLS,    // InitMoveProgressForm
      Forms;


const
  ERR_TRANSACTION       = -1;
  ERR_MSG_TRANSACTION   = 'Alignement impossible.';
  ERR_LGGEN             = -2;
  ERR_MSG_LGGEN         = 'La longueur des comptes définie dans le dossier type est supérieure à celle du dossier.';
  ERR_LGAUX             = -3;
  ERR_MSG_LGAUX         = 'La longueur des auxiliaires définie dans le dossier type est supérieure à celle du dossier.';
  ERR_LGSECTION         = -4;
  ERR_MSG_LGSECTION     = 'La longueur des sections définie dans le dossier type est supérieure à celle du dossier.';
  ERR_LGSECTIONSTD      = -5;
  ERR_MSG_LGSECTIONSTD  = 'Alignement annulé. Le paramétrage des axes du standard est incorrect.';

type
  TNatureAlignement = (alGeneraux, alTiers, alJournaux, alGuides, alSections,
                        alStdEdition,alRuptEdition,alCpteCorresp,alTabLibres,
                        alTva, alLibelleAuto);

  TInfoAligneStd = procedure ( Sender : TObject ; Erreur : integer ; Msg : string ) of object;
  TTraiteEnreg  = procedure ( T : TOB; Champ : string; var ValChamp : variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean) of object;

  TAligneRec = record
    Nature : TNatureAlignement;
    NumStd : integer;
    PP : procedure of object;
  end;

type
  TAligneStd = class
    private
      FNumStd         : integer;
      FAAligner       : array [1..20] of TAligneRec;
      FOnInformation  : TInfoAligneStd;
      FLastError      : integer;
      FLastErrorMsg   : string;
      FOkProgressForm : Boolean;
      FOkGereAna      : Boolean; // Gestion de l'analytique sur le dossier
      FResultat       : TOB;

      procedure       AlignementStandard;
      procedure       AlignementGeneraux;
      procedure       AlignementTiers;
      procedure       AlignementJournaux;
      procedure       AlignementGuides;
      procedure       AlignementSections;
      procedure       AlignementStdEdition;
      procedure       AlignementRuptEdition;
      procedure       AlignementCpteCorresp;
      procedure       AlignementTabLibres;
      procedure       AlignementTVA;
      procedure       AlignementLibelleAuto;
      procedure       AlignementQualifiantQte;
      procedure       AlignementParamSoc;

      function        StandardToDossier(TStd, TDos : TOB; PP : TTraiteEnreg) : boolean;
      procedure       TraiteChampsGeneraux ( T : TOB; Champ : string; var ValChamp : variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean);
      procedure       TraiteChampsTiers(T : TOB; Champ: string; var ValChamp: variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean);
      procedure       TraiteChampsJournaux(T : TOB; Champ: string; var ValChamp: variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean);
      procedure       TraiteChampsGuide(T : TOB; Champ: string; var ValChamp: variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean);
      procedure       TraiteChampsGuideEcr(T : TOB; Champ: string; var ValChamp: variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean);
      procedure       TraiteChampsGuideAna(T : TOB; Champ: string; var ValChamp: variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean);
      procedure       TraiteChampsSection(T : TOB; Champ: string; var ValChamp: variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean);
      procedure       TraiteChampsTVA(T : TOB; Champ: string; var ValChamp: variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean);
      procedure       TraiteChampsLibelleAuto(T : TOB; Champ: string; var ValChamp: variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean);
      procedure       TraiteChampsVentil(T : TOB; Champ: string; var ValChamp: variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean);

      function        CreationGeneraux( vInNumPlan : integer; vStNumCompte : string ) : Boolean;

      // procedure       SaveInfoListe(NumStd: integer; LIST: string);
      function        MonnaieCoherente ( NumStd : integer; DevisePrinc : string ) : boolean;
      function        LibelleNature ( Nat : TNatureAlignement ) : string;
      procedure       InsertDBSpecifGuide ( T : TOB );
      procedure       InsertDBSpecif ( T : TOB );
      //procedure       FusionneFiltre(vNumStd: integer; vTableDes, vTableSrc, vCle, vCondition: string; vEcran: TForm);
      procedure       FusionneFiltre(vNumStd : integer; vTableDes, vTableSrc, vCle, vCondition : string);
      procedure       FOnProgressFormInformation( Sender : TObject ; Erreur : integer ; Msg : string );

    public
      FRazFiltres   : Boolean;

      constructor   Create;
      destructor    Destroy; override;

      procedure     Execute;
      procedure     ChargeResultat(T: TOB);
      procedure     AjouteTAF ( alNat : TNatureAlignement; Standard : string);
      procedure     ChargeAligneStdAvecTob( vTobAligneStd : Tob);

      property      OnInformation : TInfoAligneStd read FOnInformation write FOnInformation;
      property      LastError     : integer read FLastError;
      property      LastErrorMsg  : string  read FLastErrorMsg;
      property      OkProgressForm : Boolean read FOkProgressForm write FOkProgressForm;

  end;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  uLibStdCpta;


{ TAligneStd }

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TAligneStd.ChargeResultat ( T : TOB );
{$IFDEF EAGLSERVER}
var lNumEnreg : integer;
    lTobFille : Tob;
{$ENDIF}
begin
{$IFDEF EAGLSERVER}
  // Modification de FResultat, on ne va renvoyer que le nombre d'enregistrements
  // crées pour chaque traitement, et non pas la liste détaillée
  if FResultat <> nil then
  begin
    lNumEnreg := FResultat.Detail.Count;
    FResultat.ClearDetail;
    lTobFille := Tob.Create( '', FResultat, -1);
    lTobFille.AddChampSupValeur('NATURE', 'Nombres d''enregistrements crées :');
    lTobFille.AddChampSupValeur('VALEUR', IntToStr(lNumEnreg));
    lTobFille.AddChampSupValeur('TABLE', '');
  end;
{$ENDIF}
  while FResultat.Detail.Count > 0 do
    FResultat.Detail[0].ChangeParent ( T, -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 03/04/2002
Modifié le ... :   /  /
Description .. : Fonction d'alignement du dossier par rapport à un standard
Suite ........ : pour une table donnée.
Suite ........ : Paramètres :
Suite ........ : TStd : TOB du standard
Suite ........ : TDos : TOB du dossier
Suite ........ : PP : procedure de traitement spécifique au standard
Suite ........ : Valeur retournée :  True si OK, False sinon.
Mots clefs ... : ALIGNEMENT;STANDARD;COMPTA
*****************************************************************}
function TAligneStd.StandardToDossier(TStd, TDos : TOB; PP : TTraiteEnreg) : boolean;
var i : integer;
    Prefixe, Suffixe , NomChamp : string;
    ValChamp       : variant;
    bErreur        : boolean;
    bGardeSiErreur : boolean;
begin
  Result := True;
  Prefixe := TableToPrefixe(TDos.NomTable);
  for i := 1 to TStd.NbChamps do
  begin
    NomChamp := TStd.GetNomChamp(i);
    ValChamp := TStd.GetValeur(i);
    Suffixe := ExtractSuffixe(NomChamp);
    bGardeSiErreur := False;

    // Gestion des exceptions à la règle ici
    if (Suffixe = 'PREDEFINI') or (Suffixe='NUMPLAN') then continue;
    PP ( TStd, Suffixe , ValChamp, bErreur, bGardeSiErreur);
    if bErreur then
    begin
      if not bGardeSiErreur then
        if FResultat.Detail.Count > 0 then FResultat.Detail[FResultat.Detail.Count-1].Free;

      TDos.Free;
      Result := False;
      exit;
    end
    else
    begin
      if Suffixe = 'ETABLISSEMENT' then
        TDos.PutValue( Prefixe + '_' + Suffixe, VH^.EtablisDefaut)
      else
        TDos.PutValue( Prefixe + '_' + Suffixe, ValChamp);
    end;
  end;
end;

procedure TAligneStd.AlignementGeneraux;
var   TDos, TStd, TDosInsert, T :  TOB;
      TStdv, Tv : Tob;
      i  : integer;
      iv : integer;
      Prefixe : string;
      StdlGGen : integer;
      lQueryStdv : TQuery;
begin
  // Alignement des comptes généraux
  // vérifications préliminaires
  StdlGGen := StrToInt(GetColonneSQL ('PARSOCREF','PRR_SOCDATA','PRR_NUMPLAN='+IntToStr(FNumStd)+' AND (PRR_SOCNOM="SO_LGCPTEGEN")'));
  if VH^.Cpta[fbGene].Lg < StdLgGen then
  begin
    FLastError    := ERR_LGGEN;
    FLastErrorMsg := ERR_MSG_LGGEN;
    if Assigned(FOnInformation) then FOnInformation(self,FLastError, FLastErrorMsg);
    Exit;
  end;

  TDos := TOB.Create ('',nil,-1);
  try
    if assigned (FOnInformation) then FOnInformation (self,0,'Lecture du plan de compte ...');

    TDos.LoadDetailDB ('GENERAUX','','',nil,False);
    TStd := TOB.Create ('',nil,-1);
    try
      if assigned (FOnInformation) then FOnInformation (self,0,'Lecture du dossier type ...');

      TStd.LoadDetailDB ('GENERAUXREF',IntToStr(FNumStd),'',nil,False);
      TDosInsert := TOB.Create ('',nil,-1);
      try
        Prefixe := TableToPrefixe ('GENERAUX');
        if assigned (FOnInformation) then FOnInformation (self,0,'Alignement des comptes généraux ...');

        for i:=0 to TStd.Detail.Count - 1 do
        begin
          if (TDos.FindFirst(['G_GENERAL'],[BourreEtLess(TStd.Detail[i].GetValue('GER_GENERAL'),fbGene)],False) = nil ) then
          begin
            T := TOB.Create ('GENERAUX',TDosInsert,-1);
            if not StandardToDossier(TStd.Detail[i],T , TraiteChampsGeneraux) then
              Continue;

            if FOkGereAna then
            begin
              // GCO - 19/05/2005 - Recopie des ventilations par défaut
              TStdv := TOB.Create ('', nil, -1);
              lQueryStdv := nil;
              try
                lQueryStdv := OpenSql('SELECT * FROM VENTILREF WHERE ' +
                                      'VR_NUMPLAN = ' + IntToStr(FNumStd) + ' AND ' +
                                      'VR_COMPTE = "' + BourreEtLess(TStd.Detail[i].GetValue('GER_GENERAL'), fbGene) + '" ' +
                                      'ORDER BY VR_COMPTE', True);

                TStdv.LoadDetailDB('VENTILREF', '', '', lQueryStdv,False);
                for iv := 0 to TStdv.Detail.Count - 1 do
                begin
                  Tv := TOB.Create ('VENTIL', T, -1);
                  StandardToDossier (TStdv.Detail[iv], Tv, TraiteChampsVentil);
                end;
              finally
                TStdv.Free;
                Ferme(lQueryStdv);
              end;
            end;
            // FIN GCO
          end;
        end;
        if assigned (FOnInformation) then FOnInformation (self,0,'Mise à jour de la base de données ...');

        TDosInsert.InsertDB (nil);
      finally
       TDosInsert.Free;
      end;
    finally
      TStd.Free;
    end;
  finally
    TDos.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... : 14/05/2004
Description .. : Alignement des journaux
Suite ........ : GCO - 14/05/2004 Création du Compte de Contrepartie des
Suite ........ : journaux de type Banque et Caisse si il n'existe pas
Mots clefs ... :
*****************************************************************}
procedure TAligneStd.AlignementJournaux;
var   TDos, TStd, TDosInsert, T :  TOB;
      i  : integer;
      Prefixe : string;
begin
  TDos := TOB.Create ('',nil,-1);
  try
    if assigned (FOnInformation) then FOnInformation (self,0,'Lecture des journaux ...');

    TDos.LoadDetailDB ('JOURNAL','','',nil,False);
    TStd := TOB.Create ('',nil,-1);
    try
      if assigned (FOnInformation) then FOnInformation (self,0,'Lecture du dossier type ...');
      TStd.LoadDetailDB ('JALREF',IntToStr(FNumStd),'',nil,False);
      TDosInsert := TOB.Create ('',nil,-1);
      try
        Prefixe := TableToPrefixe ('JOURNAL');
        if assigned (FOnInformation) then FOnInformation (self,0,'Alignement des journaux ...');
        for i:=0 to TStd.Detail.Count - 1 do
        begin
          if (TDos.FindFirst(['J_JOURNAL'],[TStd.Detail[i].GetValue('JR_JOURNAL')],False) = nil ) then
          begin
            T := TOB.Create ('JOURNAL',TDosInsert,-1);
            if not StandardToDossier (TStd.Detail[i], T, TraiteChampsJournaux ) then
              continue;
          end;
        end;
        if assigned (FOnInformation) then FOnInformation (self,0,'Mise à jour de la base de données ...');
        TDosInsert.InsertDB (nil);
      finally
       TDosInsert.Free;
      end;
    finally
      TStd.Free;
    end;
  finally
    TDos.Free;

    // GCO - 16/05/2005 - FQ 15424
    AvertirMultiTable('TTJOURNAL');
    // FIN GCO
  end;
end;

procedure TAligneStd.AlignementStandard;
var i : integer;
    TRes : TOB;
begin
  for i:= low (FAAligner) to high (FAAligner) do
  begin
    if (FAAligner [i].NumStd > 0) then
    begin
      FNumStd := FAAligner [i].NumStd;
      // On ne peut aligner que si la devise principale du dossier est à la même
      // que celle du dossier type considéré.
      if not MonnaieCoherente (FNumStd, GetParamSoc ('SO_DEVISEPRINC')) then
      begin
        TRes := TOB.Create ('RESULTAT', FResultat,-1);
        TRes.AddChampSupValeur('Nature',LibelleNature(FAAligner[i].Nature));
        TRes.AddChampSupValeur('Valeur','Erreur : Monnaie incohérente');
        TRes.AddChampSupValeur('Table','ERREUR');
        continue;
      end;
      FAAligner [i].PP;
      if FLastError < 0 then break;
    end;
  end;

  if FLastError >= 0 then
  begin
    AlignementParamSoc;
    AlignementQualifiantQte;
  end;
end;

procedure TAligneStd.AlignementTiers;
var   TDos, TStd, TDosInsert, T :  TOB;
      i  : integer;
      Prefixe : string;
      StdLgAux : integer;
begin
  // Alignement des comptes généraux
  // vérifications préliminaires
  StdLgAux := StrToInt(GetColonneSQL ('PARSOCREF','PRR_SOCDATA','PRR_NUMPLAN='+IntToStr(FNumStd)+' AND (PRR_SOCNOM="SO_LGCPTEAUX")'));
  if VH^.Cpta[fbAux].Lg < StdLgAux then
  begin
    FLastError    := ERR_LGAUX;
    FLastErrorMsg := ERR_MSG_LGAUX;
    if Assigned(FOnInformation) then
      FOnInformation(self,FLastError,
        FLastErrorMsg);
    exit;
  end;
  TDos := TOB.Create ('',nil,-1);
  try
    if assigned (FOnInformation) then FOnInformation (self,0,'Lecture des auxiliaires du dossier ...');
    TDos.LoadDetailDB ('TIERS','','',nil,False);
    TStd := TOB.Create ('',nil,-1);
    try
      if assigned (FOnInformation) then FOnInformation (self,0,'Lecture du dossier type ...');
      TStd.LoadDetailDB ('TIERSREF',IntToStr(FNumStd),'',nil,False);
      TDosInsert := TOB.Create ('',nil,-1);
      try
        Prefixe := TableToPrefixe ('TIERS');
        if assigned (FOnInformation) then FOnInformation (self,0,'Alignement des comptes auxiliaires ...');
        for i:=0 to TStd.Detail.Count - 1 do
        begin
          if (TDos.FindFirst(['T_AUXILIAIRE'],[BourreEtLess(TStd.Detail[i].GetValue('TRR_AUXILIAIRE'),fbAux)],False) = nil ) then
          begin
            T := TOB.Create ('TIERS',TDosInsert,-1);
            if not StandardToDossier(TStd.Detail[i],T,TraiteChampsTiers) then
              continue;
          end;
        end;
        if assigned (FOnInformation) then FOnInformation (self,0,'Mise à jour de la base de données ...');

        TDosInsert.InsertDB (nil);
      finally
       TDosInsert.Free;
      end;
    finally
      TStd.Free;
    end;
  finally
    TDos.Free;
  end;
end;

constructor TAligneStd.Create;
begin
  FOkProgressForm := False;
  FRazFiltres     := False;
  FOkGereAna      := not (GetParamSoc('SO_CPPCLSANSANA'));

  // Initialisation du résultat
  FResultat := TOB.Create ('', nil, -1);
  // Initialisation de la structure d'alignement
  FAAligner[1].Nature :=alGeneraux;
  FAAligner[1].NumStd :=0;
  FAAligner[1].PP :=AlignementGeneraux;

  FAAligner[2].Nature :=alTiers;
  FAAligner[2].NumStd :=0;
  FAAligner[2].PP :=AlignementTiers;

  FAAligner[3].Nature :=alJournaux;
  FAAligner[3].NumStd :=0;
  FAAligner[3].PP :=AlignementJournaux;

  FAAligner[4].Nature :=alGuides;
  FAAligner[4].NumStd :=0;
  FAAligner[4].PP :=AlignementGuides;

  FAAligner[5].Nature :=alSections;
  FAAligner[5].NumStd :=0;
  FAAligner[5].PP :=AlignementSections;

  FAAligner[6].Nature :=alStdEdition;
  FAAligner[6].NumStd :=0;
  FAAligner[6].PP :=AlignementStdEdition;

  FAAligner[7].Nature :=alRuptEdition;
  FAAligner[7].NumStd :=0;
  FAAligner[7].PP :=AlignementRuptEdition;

  FAAligner[8].Nature :=alCpteCorresp;
  FAAligner[8].NumStd :=0;
  FAAligner[8].PP :=AlignementCpteCorresp;

  FAAligner[9].Nature :=alTabLibres;
  FAAligner[9].NumStd :=0;
  FAAligner[9].PP :=AlignementTabLibres;

  // GCO - 04/05/2005
  // Alignement de la TVA
  FAAligner[10].Nature := alTVA;
  FAAligner[10].NumStd := 0;
  FAAligner[10].PP     := AlignementTVA;

  // Alignement des libellés AUTO
  FAAligner[11].Nature := alLibelleAuto;
  FAAligner[11].NumStd := 0;
  FAAligner[11].PP     := AlignementLibelleAuto;
  // FIN GCO

end;

destructor TAligneStd.Destroy;
begin
  FResultat.Free;
  FiniMoveProgressForm;
  inherited;
end;

procedure TAligneStd.Execute;
var lInNbTraitement, i : integer;
begin
  FLastError := 0;

  if FOkProgressForm and (not Assigned (fOnInformation)) then
  begin
    OnInformation := FOnProgressFormInformation;
    lInNbTraitement := 0;
    for i:= low (FAAligner) to high (FAAligner) do
    begin
      if (FAAligner [i].NumStd > 0) then
        Inc(lInNbTraitement);
    end;
    InitMoveProgressForm(nil, 'Alignement des Standards', 'Traitement en cours...', lInNbTraitement*4, False, True);
  end;

  // Lancement de l'alignement
  if Transactions(AlignementStandard,1) <> oeOk then
  begin
    FLastError := ERR_TRANSACTION;
    FLastErrorMsg := ERR_MSG_TRANSACTION;
    if assigned (FOnInformation) then FOnInformation (self, FLastError, FLastErrorMsg);
  end
  else
  begin
    if FLastError < 0 then
    begin
      if assigned (FOnInformation) then FOnInformation (self,0,'Erreur lors de l''alignement');
    end
    else
    begin
      SetParamSoc('SO_CPDATEDERNMAJPLAN', Now);
      if assigned (FOnInformation) then FOnInformation (self,0,'Alignement terminé.');
    end;
  end;
end;

procedure TAligneStd.TraiteChampsGeneraux(T : TOB; Champ: string; var ValChamp: variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean);
var TRes : TOB;
begin
  bErreur := False;
  if Champ = 'GENERAL' then
  begin
    ValChamp := BourreEtLess (ValChamp,fbGene);
    TRes := TOB.Create ('RESULTAT', FResultat,-1);
    TRes.AddChampSupValeur('Nature','Compte général');
    TRes.AddChampSupValeur('Valeur',ValChamp);
    TRes.AddChampSupValeur('Table','GENERAUX');
  end else
  if Champ = 'CONFIDENTIEL' then
  begin
    if ValChamp = 'X' then
      ValChamp := '1'
    else
    begin
      if (ValChamp = '-') or (ValChamp = '') or VarIsNull(ValChamp) then
        ValChamp := '0';
    end;
  end
  else
  if Pos('VENTILABLE', Champ) > 0 then
  begin
    if not FOkGereAna then
      ValChamp := '-';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TAligneStd.TraiteChampsVentil(T: TOB; Champ: string; var ValChamp: variant; var bErreur, vBoGardeSiErreur: Boolean);
begin
  // Aucun champ à traiter pour le moment
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/05/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TAligneStd.TraiteChampsTiers(T : TOB; Champ : string; var ValChamp: variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean);
var TRes : TOB;
    NatureAuxi, NatureGene : string;
begin
  bErreur := False;
  if Champ = 'AUXILIAIRE' then
  begin
    ValChamp := BourreEtLess (ValChamp, fbAux); // GCO - FQ - 13541
    TRes := TOB.Create ('RESULTAT', FResultat,-1);
    TRes.AddChampSupValeur('Nature','Compte auxiliaire');
    TRes.AddChampSupValeur('Valeur',ValChamp);
    TRes.AddChampSupValeur('Table','TIERS');
  end else
  if (Champ = 'COLLECTIF') and ( ValChamp <> '') then
  begin
    //ValChamp := BourreEtLess (ValChamp,fbAux);
    ValChamp := BourreEtLess (ValChamp, fbGene); // GCO - 11/09/2007
    NatureAuxi := T.GetValue('TRR_NATUREAUXI');
    NatureGene := GetColonneSQL ('GENERAUX','G_NATUREGENE','G_GENERAL="'+ValChamp+'"');

    if NatureGene = '' then
    begin // GCO - 11/09/2007 - FQ 21280 - Création du compte du collectif si existe pas
      bErreur := not CreationGeneraux(FNumStd, ValChamp);
    end
    else
    begin // il existe, on regarde sa nature par rapport à celle de l'auxilaire
      if ((NatureGene='COC') and (NatureAuxi='CLI')) or
         ((NatureGene='COD') and (NatureAuxi='DIV')) or
         ((NatureGene='COF') and (NatureAuxi='FOU')) or
         ((NatureGene='COS') and (NatureAuxi='SAL')) then
      else // sinon, le collectif n'est pas bon, on ne met rien
        ValChamp := '';
    end
  end else
  if Champ = 'CONFIDENTIEL' then
  begin
    if ValChamp = 'X' then
      ValChamp := '1'
    else
    begin
      if (ValChamp = '-') or (ValChamp = '') or VarIsNull(ValChamp) then
        ValChamp := '0';
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... : 12/05/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TAligneStd.TraiteChampsJournaux(T : TOB; Champ: string; var ValChamp: variant; var bErreur : boolean; var vBoGardeSiErreur : Boolean);
var TRes : TOB;
    TSoucheCreation : Tob;
    lStLibelle      : string;
    lStAbrege       : string;
    lStNatureGene   : string;
    lQuery          : TQuery;
begin
  bErreur := False;
  TRes := nil;
  if Champ = 'JOURNAL' then
  begin
    TRes := TOB.Create ('RESULTAT', FResultat,-1);
    TRes.AddChampSupValeur('Nature','Journal');
    TRes.AddChampSupValeur('Valeur',ValChamp);
    TRes.AddChampSupValeur('Table','JOURNAL');
  end
  else
  if Champ = 'COMPTEURNORMAL' then
  begin
    if ValChamp <> '' then
    begin
      if not ExisteSQL('SELECT SH_SOUCHE FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+ValChamp+'"') then
      begin
        // FQ 10195 - Création de la souche si elle n'existe pas
        // GCO - 19/11/2004 - FQ 15002
        if ValChamp = 'CPT' then
        begin
          // Création de la Souche Comptable
          lStLibelle := 'Souche Comptable';
          lStAbrege  := lStLibelle;
        end
        else
        begin
          lQuery := OpenSql('SELECT J_LIBELLE, J_ABREGE FROM JOURNAL WHERE ' +
                    'J_JOURNAL = "' + ValChamp + '" ORDER BY J_JOURNAL', True);
          if not lQuery.Eof then
          begin
            lStLibelle := lQuery.FindField('J_LIBELLE').AsString;
            lStAbrege  := lQuery.FindField('J_ABREGE').AsString;
           end;
           Ferme( lQuery );
        end;

        TSoucheCreation := Tob.Create('SOUCHE', nil, -1 );
        TSoucheCreation.PutValue('SH_TYPE', 'CPT');
        TSoucheCreation.PutValue('SH_SOUCHE', ValChamp);
        TSoucheCreation.PutValue('SH_LIBELLE', lStLibelle);
        TSoucheCreation.PutValue('SH_ABREGE', lStAbrege);
        TSoucheCreation.PutValue('SH_SOCIETE', V_Pgi.CodeSociete);
        TSoucheCreation.PutValue('SH_NUMDEPART',  1);
        TSoucheCreation.PutValue('SH_NUMDEPARTS', 1);
        TSoucheCreation.PutValue('SH_NUMDEPARTP', 1);

        // GCO - 19/11/2004 - FQ 15003
        bErreur := not TSoucheCreation.InsertDB( nil, False );
        TSoucheCreation.Free;
      end;
    end;
  end
  else
  if Champ = 'COMPTEURSIMUL' then
  begin
    if ValChamp <> '' then
    begin
      if not ExisteSQL('SELECT SH_SOUCHE FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE = "SIM"') then
      begin
        // FQ 10195 - Création de la souche SIM si elle n'existe pas
        TSoucheCreation := Tob.Create('SOUCHE', nil, -1 );
        TSoucheCreation.PutValue('SH_TYPE', 'CPT');
        TSoucheCreation.PutValue('SH_SOUCHE', 'SIM');
        TSoucheCreation.PutValue('SH_LIBELLE', 'Simulation');
        TSoucheCreation.PutValue('SH_ABREGE', 'Souche multi jour');
        TSoucheCreation.PutValue('SH_SOCIETE', V_Pgi.CodeSociete);
        TSoucheCreation.PutValue('SH_SIMULATION', 'X');
        TSoucheCreation.PutValue('SH_NUMDEPART',  1);
        TSoucheCreation.PutValue('SH_NUMDEPARTS', 1);
        TSoucheCreation.PutValue('SH_NUMDEPARTP', 1);

        // GCO - 19/11/2004 - FQ 15003
        bErreur := not TSoucheCreation.InsertDB( nil, False );
        TSoucheCreation.Free;
      end;
    end;
  end
  else
  if Champ = 'CONTREPARTIE' then
  begin
    if ValChamp <> '' then
    begin
      ValChamp := BourreEtLess(ValChamp,fbGene);
      if not Presence ('GENERAUX','G_GENERAL', ValChamp) then
      begin
        // GCO - FQ 13507 - Création du compte de contrepartie
        // GCO - 30/11/2004 - FQ 15003
        bErreur  := not CreationGeneraux(T.GetValue('JR_NUMPLAN'),T.GetValue('JR_CONTREPARTIE'));
      end
      else
      begin // FQ 13507 - 02/05/2005 - Si Journal = BQE ou CAI, si le compte de
            // contrepartie existe, tester que G_NATUREGENE = JR_NATUREJAL
        if (T.GetString('JR_NATUREJAL') = 'BQE') or
           (T.GetString('JR_NATUREJAL') = 'CAI') then
        begin
          lStNatureGene := GetColonneSql( 'GENERAUX', 'G_NATUREGENE', 'G_GENERAL = "' + ValChamp + '"');
          if lStNatureGene <> T.GetString('JR_NATUREJAL') then
          begin
            bErreur := True;
            vBoGardeSiErreur := True;
            if TRes <> nil then
            begin
              TRes.SetString('Nature', 'Journal ' + T.GetString('JR_JOURNAL') + ' non aligné.');
              TRes.SetString('Valeur', 'Erreur avec la nature du compte ' + ValChamp);
              TRes.SetString('Table', 'JOURNAL');
            end
            else
            begin
              TRes := TOB.Create ('RESULTAT', FResultat,-1);
              TRes.AddChampSupValeur('Nature', 'Journal ' + T.GetString('JR_JOURNAL') + ' non aligné.');
              TRes.AddchampSupValeur('Valeur', 'Erreur avec la nature du compte ' + ValChamp);
              TRes.AddChampSupValeur('Table', 'JOURNAL');
            end;
          end;
        end;
      end;
    end;
  end
  else
  begin
    if Copy(Champ,1,9)='CPTEREGUL' then
    begin
      if ValChamp <> '' then
      begin
        ValChamp := BourreEtLess(ValChamp,fbGene);
        if not Presence ('GENERAUX','G_GENERAL', ValChamp) then
          ValChamp := '';
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... : 12/05/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TAligneStd.AlignementGuides;
var TDos, TStd,TStde, TStda, TDosInsert, T, Te, Ta :  TOB;
    i ,ie, ia : integer;
    lInAdd : integer;
    lStOldKey : string;
    Prefixe : string;

    function  _CalculCleGuide( vStGU_Type : string; vInAdd : integer = 0) : string;
    var lQuery : TQuery;
        lSt : string;
    begin
      lQuery := nil;
      try
        try
          lQuery := OpenSQL('SELECT MAX(GU_GUIDE) LEMAX FROM GUIDE WHERE GU_TYPE = "' +
                    vStGU_Type + '" ORDER BY 1',True);

          lSt := lQuery.FindField('LEMAX').AsString ;
          if lSt = '' then
            lSt := '000' ;

          lSt := Format('%3.3d', [StrToInt( lSt ) + 1 + vInAdd]);

        except
          on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : CalculCleGuide');
        end;
      finally
        Ferme( lQuery ) ;
        Result := lSt;
      end;
   end;

begin
  // Alignement des comptes généraux
  TDos := TOB.Create ('',nil,-1);
  try
    if assigned (FOnInformation) then FOnInformation (self,0,'Chargement des guides ...');
    TDos.LoadDetailDB ('GUIDE','','',nil,False);
    TStd := TOB.Create ('',nil,-1);
    try
      if assigned (FOnInformation) then FOnInformation (self,0,'Lecture du dossier type ...');
      TStd.LoadDetailDB ('GUIDEREF',IntToStr(FNumStd),'',nil,False);
      TDosInsert := TOB.Create ('',nil,-1);
      try
        Prefixe := TableToPrefixe ('GUIDE');
        if assigned (FOnInformation) then FOnInformation (self,0,'Alignement des guides ...');

        lInAdd := 0;
        for i:=0 to TStd.Detail.Count - 1 do
        begin
          // GCO - 08/12/2004 - FQ 15003
          // Sauvegarde de l'ancienne pour Clé afin de charger correctement ECRGUIREF
          lStOldKey := TStd.Detail[i].GetValue('GDR_GUIDE');

          //// FQ 10685
          //if (TDos.FindFirst(['GU_TYPE','GU_GUIDE'],[TStd.Detail[i].GetValue('GDR_TYPE'),TStd.Detail[i].GetValue('GDR_GUIDE')],False) <> nil ) then
          //begin
            // Calcul de la clé du guide que l'on va ajouté
          TStd.Detail[i].PutValue('GDR_GUIDE', _CalculCleGuide( TStd.Detail[i].GetValue('GDR_TYPE'), lInAdd));
          Inc( lInAdd );
          //end;

          T := TOB.Create ('GUIDE',TDosInsert,-1);
          // GCO - 27/10/2004 - FQ 13042
          T.PutValue('GU_ETABLISSEMENT', VH^.EtablisDefaut);

          if not StandardToDossier(TStd.Detail[i],T , TraiteChampsGuide) then continue;
          TStde := TOB.Create ('',nil,-1);
          try
            TStde.LoadDetailDB('ECRGUIREF',IntToStr(TStd.Detail[i].GetValue('GDR_NUMPLAN'))+';"'+TStd.Detail[i].GetValue('GDR_TYPE')+'";"'+lStOldKey+'"','',nil,False);
            for ie := 0 to TStde.Detail.Count - 1 do
            begin
              Te := TOB.Create ('ECRGUI',T,-1);
              // Enregistrement de l'ancienne ou nouvelle clé dans EGR_GUIDE
              // GCO - 08/12/2004 - FQ 15003
              TStde.Detail[ie].PutValue('EGR_GUIDE', TStd.Detail[i].GetValue('GDR_GUIDE'));
              StandardToDossier (TStde.Detail[ie], Te, TraiteChampsGuideEcr);
              TStda := TOB.Create ('',nil,-1);
              try
                TStda.LoadDetailDB('ANAGUIREF',IntToStr(TStd.Detail[i].GetValue('GDR_NUMPLAN'))+';"'+TStd.Detail[i].GetValue('GDR_TYPE')+'";"'+ lStOldKey +'";'+IntToStr(TStde.Detail[ie].GetValue('EGR_NUMLIGNE')),'',nil,False);
                for ia := 0 to TStda.Detail.Count - 1 do
                begin
                  Ta := TOB.Create ('ANAGUI',T,-1);
                  // Enregistrement de l'ancienne ou nouvelle clé dans AGR_GUIDE
                  // GCO - 08/12/2004 - FQ 15003
                  TStda.Detail[ia].PutValue('AGR_GUIDE', TStd.Detail[i].GetValue('GDR_GUIDE'));
                  StandardToDossier (TStda.Detail[ia], Ta, TraiteChampsGuideAna);
                end;
              finally
                TStda.Free;
              end;
            end;
          finally
            TStde.Free;
          end;
        end;

        if assigned (FOnInformation) then  FOnInformation (self,0,'Mise à jour de la base de données ...');
//        TDosInsert.InsertDB (nil);
// Temporairement, on fait une insertion ' à l'ancienne' dans la base pour contourner
// le problème des accolades qui sont dégagées par le changeSQL
          InsertDBSpecifGuide ( TDosInsert );
      finally
       TDosInsert.Free;
      end;
    finally
      TStd.Free;
    end;
  finally
    TDos.Free;
  end;
end;

procedure TAligneStd.TraiteChampsGuide(T: TOB; Champ: string; var ValChamp: variant; var bErreur: boolean; var vBoGardeSiErreur : Boolean);
var TRes : TOB;
begin
  bErreur := False;
  if Champ = 'GUIDE' then
  begin
    TRes := TOB.Create ('RESULTAT', FResultat,-1);
    TRes.AddChampSupValeur('Nature','Guide');
    TRes.AddChampSupValeur('Valeur',ValChamp);
    TRes.AddChampSupValeur('Table','GUIDE');
  end
  else if Champ = 'JOURNAL' then
  begin
  if ValChamp <> '' then
    bErreur := not ExisteSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_JOURNAL="'+ValChamp+'"');
  end ;
end;

procedure TAligneStd.TraiteChampsGuideEcr(T: TOB; Champ: string; var ValChamp: variant; var bErreur: boolean; var vBoGardeSiErreur : Boolean);
begin
  bErreur := False;
  if Champ = 'GENERAL' then
  begin
    if ValChamp <> '' then
    begin
      ValChamp := BourreEtLess(ValChamp,fbGene);
      if not Presence ('GENERAUX','G_GENERAL', ValChamp) then
      begin
        ValChamp := '';
        // on rend la zone saisissable
        T.PutValue('EGR_ARRET','-'+Copy(T.GetValue('EGR_ARRET'),2,Length(T.GetValue('EGR_ARRET'))-1));
      end;
    end;
  end else
  if Champ = 'AUXILIAIRE' then
  begin
    if ValChamp <> '' then
    begin
      ValChamp := BourreEtLess(ValChamp,fbAux);
      if not Presence ('TIERS','T_AUXILIAIRE', ValChamp) then
      begin
        ValChamp := '';
        // on rend la zone saisissable
        T.PutValue('EGR_ARRET',Copy(T.GetValue('EGR_ARRET'),1,1)+'-'+Copy(T.GetValue('EGR_ARRET'),3,Length(T.GetValue('EGR_ARRET'))-2));
      end;
    end;
  end;
end;

procedure TAligneStd.TraiteChampsGuideAna(T: TOB; Champ: string; var ValChamp: variant; var bErreur: boolean; var vBoGardeSiErreur : Boolean);
begin
  bErreur := False;
  if Champ = 'SECTION' then
  begin
    if ValChamp <> '' then
    begin
      ValChamp := BourreEtLess(ValChamp,AxeToFb (T.GetValue('AGR_AXE')));
      if not Presence ('SECTION','S_SECTION', ValChamp) then
      begin
        ValChamp := '';
        // on rend la zone saisissable
        T.PutValue('EGR_ARRET','-'+Copy(T.GetValue('AGR_ARRET'),2,Length(T.GetValue('AGR_ARRET'))-1));
      end;
    end;
  end;
end;

procedure TAligneStd.AlignementSections;
var   TDos, TStd, TDosInsert, T :  TOB;
      i  : integer;
      Prefixe : string;
      StdLgSection1 : integer;
      StdLgSection2 : integer;
      StdLgSection3 : integer;
      StdLgSection4 : integer;
      StdLgSection5 : integer;
      lStAxeTvaStd : string;
      lStLgAxeRef1 : string;
      lStLgAxeRef2 : string;
      lStLgAxeRef3 : string;
      lStLgAxeRef4 : string;
      lStLgAxeRef5 : string;

begin
  // GCO - 20/09/2007 - FQ 20453
  StdLgSection1 := 0;
  StdLgSection2 := 0;
  StdLgSection3 := 0;
  StdLgSection4 := 0;
  StdLgSection5 := 0;

  lStLgAxeRef1 := GetColonneSQL ('AXEREF','XRE_LONGSECTION','XRE_NUMPLAN='+IntToStr(FNumStd)+' AND XRE_AXE="A1"');
  if lStLgAxeRef1 <> '' then
    StdLgSection1 := StrToInt(GetColonneSQL ('AXEREF','XRE_LONGSECTION','XRE_NUMPLAN='+IntToStr(FNumStd)+' AND XRE_AXE="A1"'));

  lStLgAxeRef2 := GetColonneSQL ('AXEREF','XRE_LONGSECTION','XRE_NUMPLAN='+IntToStr(FNumStd)+' AND XRE_AXE="A2"');
  if lStLgAxeRef2 <> '' then
    StdLgSection2 := StrToInt(GetColonneSQL ('AXEREF','XRE_LONGSECTION','XRE_NUMPLAN='+IntToStr(FNumStd)+' AND XRE_AXE="A2"'));

  lStLgAxeRef3 := GetColonneSQL ('AXEREF','XRE_LONGSECTION','XRE_NUMPLAN='+IntToStr(FNumStd)+' AND XRE_AXE="A3"');
  if lStLgAxeRef3 <> '' then
    StdLgSection3 := StrToInt(GetColonneSQL ('AXEREF','XRE_LONGSECTION','XRE_NUMPLAN='+IntToStr(FNumStd)+' AND XRE_AXE="A3"'));

  lStLgAxeRef4 := GetColonneSQL ('AXEREF','XRE_LONGSECTION','XRE_NUMPLAN='+IntToStr(FNumStd)+' AND XRE_AXE="A4"');
  if lStLgAxeRef4 <> '' then
    StdLgSection4 := StrToInt(GetColonneSQL ('AXEREF','XRE_LONGSECTION','XRE_NUMPLAN='+IntToStr(FNumStd)+' AND XRE_AXE="A4"'));

  lStLgAxeRef5 := GetColonneSQL ('AXEREF','XRE_LONGSECTION','XRE_NUMPLAN='+IntToStr(FNumStd)+' AND XRE_AXE="A5"');
  if lStLgAxeRef5 <> '' then
    StdLgSection5 := StrToInt(GetColonneSQL ('AXEREF','XRE_LONGSECTION','XRE_NUMPLAN='+IntToStr(FNumStd)+' AND XRE_AXE="A5"'));

  if (lStLgAxeRef1 = '') or (lStLgAxeRef2 = '') or (lStLgAxeRef3 = '') or
     (lStLgAxeRef4 = '') or (lStLgAxeRef5 = '') then
  begin
    FLastError := ERR_LGSECTIONSTD;
    FLastErrorMsg := ERR_MSG_LGSECTIONSTD;
    if Assigned(FOnInformation) then FOnInformation(self,FLastError, FLastErrorMsg);
    Exit;
  end;  
  // FIN GCO - 20/09/2007

  if (VH^.Cpta[fbAxe1].Lg < StdLgSection1) or
     (VH^.Cpta[fbAxe2].Lg < StdLgSection2) or
     (VH^.Cpta[fbAxe3].Lg < StdLgSection3) or
     (VH^.Cpta[fbAxe4].Lg < StdLgSection4) or
     (VH^.Cpta[fbAxe5].Lg < StdLgSection5) then
  begin
    FLastError := ERR_LGSECTION;
    FLastErrorMsg := ERR_MSG_LGSECTION;
    if Assigned(FOnInformation) then FOnInformation(self,FLastError, FLastErrorMsg);
    Exit;
  end;
  // FIN GCO

  lStAxeTvaStd := GetColonneSql('PARSOCREF', 'PRR_SOCDATA',
                  'PRR_NUMPLAN = ' + IntToStr(FNumStd) + ' AND ' +
                  'PRR_SOCNOM = "SO_CPPCLAXETVA"');

  // GCO - 19/10/2007 - FQ 20935
  if GetParamSocSecur('SO_CPPCLSAISIETVA', False, True) then
    CreationSectionTVA('');
  
  // Si le dossier ne gère pas l'analytique, inutile d'aligner sur le standard
  // vérifications préliminaires
  TDos := TOB.Create ('',nil,-1);
  try
    if assigned (FOnInformation) then
      FOnInformation (self,0,'Lecture des sections du dossier ...');
    TDos.LoadDetailDB ('SECTION','','',nil,False);
    TStd := TOB.Create ('',nil,-1);
    try
      if assigned (FOnInformation) then
        FOnInformation (self,0,'Lecture du dossier type ...');
      TStd.LoadDetailDB ('SECTIONREF',IntToStr(FNumStd),'',nil,False);
      TDosInsert := TOB.Create ('',nil,-1);
      try
        Prefixe := TableToPrefixe ('SECTION');
        if assigned (FOnInformation) then FOnInformation (self,0,'Alignement des sections ...');

        for i:=0 to TStd.Detail.Count - 1 do
        begin
          // GCO - Pas de mise à jour des sections TVA de cette façon, car l'axe
          // TVA du standard et celui du dossier peuvent être différent.
          if (TStd.Detail[i].GetString('SRE_AXE') = lStAxeTvaStd) then Continue;

          if (TDos.FindFirst(['S_SECTION'],[BourreEtLess(TStd.Detail[i].GetValue('SRE_SECTION'),fbSect)],False) = nil ) then
          begin
            T := TOB.Create ('SECTION',TDosInsert,-1);
            if not StandardToDossier(TStd.Detail[i],T,TraiteChampsSection) then
              continue;
          end;
        end;

        if assigned (FOnInformation) then FOnInformation (self,0,'Mise à jour de la base de données ...');
        TDosInsert.InsertDB (nil);
      finally
       TDosInsert.Free;
      end;
    finally
      TStd.Free;
    end;
  finally
    TDos.Free;
  end;
  // Alignement des lois de répartition
  YFSChargeTableDepuisStandard('CLOIVENTIL',FNumStd,'');
  // Recopie des ventilations associées aux lois
  TStd := TOB.Create ('',nil,-1);
  try
    TStd.LoadDetailDBFromSQL('VENTILREF','SELECT * FROM VENTILREF WHERE ' +
                            'VR_NATURE="CLV" AND VR_NUMPLAN="'+IntToStr(FNumStd)+'" ');
    if TStd.Detail.Count > 0 then
    begin
      T := TOB.Create('',nil,-1);
      try
        for i := 0 to TStd.Detail.Count - 1 do
        begin
          TDos := TOB.Create ('VENTIL', T, -1);
          StandardToDossier (TStd.Detail[i], TDos, TraiteChampsVentil);
        end;
        T.InsertOrUpdateDB;
      finally
        T.Free;
      end;
    end;
  finally
    TStd.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TAligneStd.TraiteChampsSection(T: TOB; Champ: string; var ValChamp: variant; var bErreur: boolean; var vBoGardeSiErreur : Boolean);
begin
  bErreur := False;

  // GCO - 10/10/2005 - FQ 16508
  if Champ = 'SECTION' then
  begin
    ValChamp := BourreEtLess (ValChamp, AxeToFb(T.GetString('SRE_AXE')));
  end;
  // FIN - GCO

  if Champ = 'CONFIDENTIEL' then
  begin
    if ValChamp = 'X' then
      ValChamp := '1'
    else
    begin
      if (ValChamp = '-') or (ValChamp = '') or VarIsNull(ValChamp) then
        ValChamp := '0';
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/05/2005
Modifié le ... :   /  /
Description .. : Alignement de la TVa ( TABLE TXCPTTVA )
Mots clefs ... :
*****************************************************************}
procedure TAligneStd.AlignementTVA;
var TDos, TStd, TDosInsert, T :  TOB;
    i  : integer;
    Prefixe : string;

    procedure _AlignementChoixCod;
    begin
      LoadStandardMAJ(FNumStd,'CHOIXCOD', 'CHOIXCODREF', 'TYPE;CODE', ' AND CCR_TYPE = "TX1"');
      LoadStandardMAJ(FNumStd,'CHOIXCOD', 'CHOIXCODREF', 'TYPE;CODE', ' AND CCR_TYPE = "TX2"');
    end;

begin
  TDos := TOB.Create ('',nil,-1);
  try
    if Assigned (FOnInformation) then FOnInformation (self,0,'Lecture des taux de TVA ...');
    TDos.LoadDetailDB ('TXCPTTVA','','',nil,False);
    TStd := TOB.Create ('',nil,-1);
    try
      if Assigned (FOnInformation) then FOnInformation (self,0,'Lecture du dossier type ...');
      TStd.LoadDetailDB ('TXCPTTVAREF',IntToStr(FNumStd),'',nil,False);
      TDosInsert := TOB.Create ('',nil,-1);
      try
        Prefixe := TableToPrefixe ('TXCPTTVA');
        if Assigned (FOnInformation) then FOnInformation (self,0,'Alignement des taux de TVA ...');
        for i:=0 to TStd.Detail.Count - 1 do
        begin
          // Ecrasement systématique des Taux de TVA même si il existe deja dans le dossier
          T := TOB.Create ('TXCPTTVA',TDosInsert,-1);
          if not StandardToDossier(TStd.Detail[i],T , TraiteChampsTVA) then
            Continue;
        end;
        if Assigned (FOnInformation) then FOnInformation (self,0,'Mise à jour de la base de données ...');
        TDosInsert.InsertDB (nil);
        _AlignementChoixCod;
      finally
       TDosInsert.Free;
      end;
    finally
      TStd.Free;
    end;
  finally
    TDos.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TAligneStd.TraiteChampsTVA(T: TOB; Champ: string; var ValChamp: variant; var bErreur, vBoGardeSiErreur: Boolean);
var TRes : TOB;
begin
  bErreur := False;
  // Traitement ici si cas particulier à gérer
  if Champ = 'CODETAUX' then
  begin
    if ExisteSQL('SELECT TV_TVAOUTPF FROM TXCPTTVA WHERE ' +
                 'TV_TVAOUTPF = "' + T.GetString('TVR_TVAOUTPF') + '" AND ' +
                 'TV_CODETAUX = "' + ValChamp + '" AND ' +
                 'TV_REGIME   = "' + T.GetString('TVR_REGIME') + '"') then
    begin
      // Suppression du Code TVA Existant
      ExecuteSQL('DELETE FROM TXCPTTVA WHERE ' +
                 'TV_TVAOUTPF = "' + T.GetString('TVR_TVAOUTPF') + '" AND ' +
                 'TV_CODETAUX = "' + ValChamp + '" AND ' +
                 'TV_REGIME   = "' + T.GetString('TVR_REGIME') + '"');
    end;

    TRes := TOB.Create ('RESULTAT', FResultat,-1);
    TRes.AddChampSupValeur('Nature', 'Tva - Code ' + T.GetString('TVR_TVAOUTPF'));
    TRes.AddChampSupValeur('Valeur', 'Code TVA/TPF ' + ValChamp + ' - Régime fiscal ' + T.GetString('TVR_REGIME'));
    TRes.AddChampSupValeur('Table', 'TXCPTTVA');
  end
  else
  if (Champ = 'SOCIETE') then
    ValChamp := V_Pgi.CodeSociete
  else
  begin
    if (Champ = 'CPTEACH') or (Champ = 'CPTEVTE') or
       (Champ = 'CPTVTERG') or (Champ = 'CPTACHRG') then
    begin
      if (ValChamp <> '') and (not Presence ('GENERAUX','G_GENERAL', BourreEtLess( ValChamp, fbGene))) then
        bErreur := not CreationGeneraux( T.GetInteger('TVR_NUMPLAN'), ValChamp);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TAligneStd.AlignementLibelleAuto;
var TDos, TStd, TDosInsert, T :  TOB;
    i  : integer;
    Prefixe : string;
begin
  TDos := TOB.Create ('',nil,-1);
  try
    if Assigned (FOnInformation) then FOnInformation (self,0,'Lecture des libellés automatiques ...');
    TDos.LoadDetailDB ('REFAUTO', '', '', nil, False);
    TStd := TOB.Create ('', nil, -1);
    try
      if Assigned (FOnInformation) then FOnInformation (self,0,'Lecture du dossier type ...');
      TStd.LoadDetailDB ('REFAUTOREF', IntToStr(FNumStd), '', nil, False);
      TDosInsert := TOB.Create ('', nil, -1);
      try
        Prefixe := TableToPrefixe ('REFAUTO');
        if Assigned (FOnInformation) then FOnInformation (self,0,'Alignement des libellés automatiques ...');
        for i:=0 to TStd.Detail.Count - 1 do
        begin
          // Ecrasement systématique des Libellé Auto même si il existe deja dans le dossier
          T := TOB.Create ('REFAUTO', TDosInsert, -1);
          if not StandardToDossier(TStd.Detail[i], T, TraiteChampsLibelleAuto) then
            Continue;
        end;
        if Assigned (FOnInformation) then FOnInformation (self,0,'Mise à jour de la base de données ...');
        TDosInsert.InsertDB (nil);
      finally
       TDosInsert.Free;
      end;
    finally
      TStd.Free;
    end;
  finally
    TDos.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TAligneStd.TraiteChampsLibelleAuto(T: TOB; Champ: string; var ValChamp: variant; var bErreur, vBoGardeSiErreur: Boolean);
var TRes : TOB;
begin
  bErreur := False;
  // Traitement ici si cas particulier à gérer
  if Champ = 'CODE' then
  begin
    if not Presence ('REFAUTO','RA_CODE', ValChamp) then
    begin
      TRes := TOB.Create ('RESULTAT', FResultat,-1);
      TRes.AddChampSupValeur('Nature', 'Libellé automatique');
      TRes.AddChampSupValeur('Valeur', ValChamp + ' ' + T.GetString('RAF_LIBELLE'));
      TRes.AddChampSupValeur('Table', 'REFAUTO');
    end
    else
      bErreur := True; // On conserve l'enregistrement deja existant
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/04/2006
Modifié le ... :   /  /
Description .. : FQ 17879 - On aligne tous les qualifiants Quantité.
Mots clefs ... :
*****************************************************************}
procedure TAligneStd.AlignementQualifiantQte;
begin
  LoadStandardMaj(FNumStd,'CHOIXCOD', 'CHOIXCODREF', 'TYPE;CODE', ' AND CCR_TYPE = "QME"');
  AvertirTable('TTQUALUNITMESURE');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/04/2006
Modifié le ... : 17/10/2007    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TAligneStd.AlignementParamSoc;
var lSt : string;
begin
  // GCO - 17/10/2007 - FQ 20935 - plus d'activation auto dans le dossier
  (*
  if GetParamSocSecur('SO_CPPCLSAISIETVA', False, True ) = False then
  begin
    if (GetColonneSql('PARSOCREF', 'PRR_SOCDATA', ' PRR_NUMPLAN = ' + IntToStr(FNumStd) +
                      ' AND PRR_SOCNOM = "SO_CPPCLSAISIETVA"') = 'X') then

    begin // On activate la gestion de TVa dans l'analytique
      ActivationTvaSurDossier( True, FNumStd );
    end;
  end; *)

  lSt := GetColonneSql('PARSOCREF', 'PRR_SOCDATA', 'PRR_NUMPLAN = ' + IntToStr(FNumStd) +
                   ' AND PRR_SOCNOM = "SO_CPPCLSAISIEQTE"');

  if lSt = 'X' then
    SetParamSoc('SO_CPPCLSAISIEQTE', True)
  else
    SetParamSoc('SO_CPPCLSAISIEQTE', False);
end;

procedure TAligneStd.AjouteTAF(alNat: TNatureAlignement; Standard: string);
var i : integer;
begin
  if Standard = '' then exit;
  for i:=Low (FAAligner) to High (FAAligner) do
  begin
    if FAAligner[i].Nature = alNat then
    begin
      FAAligner[i].NumStd := StrToInt ( Standard );
      break;
    end;
  end;
end;

procedure TAligneStd.AlignementStdEdition;
begin
  if FRazFiltres then
    LoadStandardMaj(FNumStd,'FILTRES','FILTRESREF', 'TABLE;LIBELLE', '', FRazFiltres)
  else
    FusionneFiltre(FNumStd, 'FILTRES', 'FILTRESREF', 'TABLE;LIBELLE', '');

  //SaveInfoListe(FNumStd, 'MULMMVTS');
  //SaveInfoListe(FNumStd, 'MULMANAL');
  //SaveInfoListe(FNumStd, 'MULVMVTS');
end;

(*
procedure TAligneStd.SaveInfoListe (NumStd : integer; LIST : string);
var  Q1, Q2 : TQuery;
begin
       Q1 := OpenSQL('SELECT * FROM LISTEREF WHERE LIR_NUMPLAN='+IntToStr (FNumStd)+ ' AND LIR_LISTE="'+LIST+'" '
       + 'AND LIR_UTILISATEUR="---"', TRUE);
  if not Q1.Eof then
  begin
       Q2 := OpenSQL('SELECT * FROM LISTE WHERE LI_LISTE="'+LIST+'" '
       + 'AND LI_UTILISATEUR="---"', FALSE);
       if Q2.Eof then
       begin
            Q2.Insert;
            InitNew(Q2);
            Q2.FindField('LI_NUMPLAN').AsInteger := NumStd;
            Q2.FindField('LI_LISTE').AsString := LIST;
            Q2.FindField('LI_UTILISATEUR').AsString := '---';
            Q2.FindField('LI_LIBELLE').AsString := Q1.findField ('LIR_LIBELLE').asstring;
            Q2.FindField('LI_SOCIETE').AsString := Q1.findField ('LIR_SOCIETE').asstring;
            Q2.FindField('LI_NUMOK').AsVariant := Q1.findField ('LIR_NUMOK').AsVariant;
            Q2.FindField('LI_TRIOK').AsVariant := Q1.findField ('LIR_TRIOK').AsVariant;
            Q2.FindField('LI_LANGUE').AsString := Q1.findField ('LIR_LANGUE').asstring;
            Q2.FindField('LI_DATA').AsVariant := Q1.findField ('LIR_DATA').AsVariant;
            Q2.Post;
       end
       else
           Q2.Edit;
           Q2.FindField('LI_DATA').Asstring := Q1.findField ('LIR_DATA').Asstring;
           Q2.Post;
       begin
       end;
       Ferme (Q2);
  end;
  Ferme (Q1);
end; *)

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TAligneStd.AlignementCpteCorresp;
var lQuery : TQuery;
begin
  // GCO - 12/09/2007 - FQ 21297
  lQuery := nil;
  try
    lQuery := OpenSql('SELECT PRR_SOCNOM, PRR_SOCDATA FROM PARSOCREF WHERE ' +
                      'PRR_NUMPLAN = ' + IntToStr(FNumStd) + ' AND ' +
                      'PRR_SOCNOM LIKE "SO_CORS%" ORDER BY PRR_SOCNOM', True);

    while not lQuery.Eof do
    begin
      if lQuery.FindField('PRR_SOCDATA').AsString = 'X' then
      begin // Plan de corresp coché dans le standard, on coche dans le dossier
        SetParamSoc( lQuery.FindField('PRR_SOCNOM').AsString, 'X');
      end;
      lQuery.Next;
    end;

  finally
    Ferme(lQuery);
  end;

  LoadStandardMaj(FNumStd,'CORRESP','CORRESPREF','TYPE;CORRESP', ' ');
end;

////////////////////////////////////////////////////////////////////////////////

procedure TAligneStd.AlignementRuptEdition;
begin
  // GCO - Condition where à ajouter pour ne prendre que les ruptures d'éditions
  LoadStandardMaj(FNumStd,'CHOIXCOD','CHOIXCODREF','TYPE;CODE', ' ');
  LoadStandardMaj(FNumStd,'RUPTURE','RUPTUREREF','NATURERUPT;PLANRUPT;CLASSE', ' ');
end;

procedure TAligneStd.AlignementTabLibres;
begin
  LoadStandardMaj(FNumStd,'CHOIXCOD','CHOIXCODREF','TYPE;CODE', ' AND CCR_TYPE="NAT"');
  LoadStandardMaj(FNumStd,'NATCPTE','NATCPTEREF','TYPECPTE;NATURE', ' ');
end;

function TAligneStd.MonnaieCoherente(NumStd: integer; DevisePrinc: string): boolean;
begin
  Result := (DevisePrinc = GetColonneSQL ('PARSOCREF','PRR_SOCDATA',
            'PRR_SOCNOM="SO_DEVISEPRINC" AND PRR_NUMPLAN='+IntToStr(NumStd)));
end;

function TAligneStd.LibelleNature(Nat: TNatureAlignement): string;
begin
  case Nat of
    alGeneraux    : Result := 'Généraux';
    alTiers       : Result := 'Tiers';
    alJournaux    : Result := 'Journaux';
    alGuides      : Result := 'Guides';
    alSections    : Result := 'Sections';
    alStdEdition  : Result := 'Standards d''éditions';
    alRuptEdition : Result := 'Plans de rupture des éditions';
    alCpteCorresp : Result := 'Comptes de correspondance';
    alTabLibres   : Result := 'Tables libres';
    alTVA         : Result := 'TVA';
    alLibelleAuto : Result := 'Libellé Auto';
  end;
  Result := TraduireMemoire (Result);
end;

procedure TAligneStd.InsertDBSpecifGuide(T: TOB);
var i,j,k : integer;
    TGuide, TEcrGuide, TAnaGuide : TOB;
begin
  for i:=0 to T.Detail.Count - 1 do
  begin
    TGuide := T.Detail[i];
    TGuide.InsertDBByNivel(False,1,1);
    for j := 0 to TGuide.Detail.Count - 1 do
    begin
      TEcrGuide := TGuide.Detail[j];
      InsertDBSpecif ( TEcrGuide );
      for k := 0 to TEcrGuide.Detail.Count - 1 do
      begin
        TAnaGuide := TEcrGuide.Detail[k];
        InsertDBSpecif ( TAnaGuide );
      end;
    end;
  end;
end;

procedure TAligneStd.InsertDBSpecif(T: TOB);
var   i : integer;
      Q : TQuery;
      stWhere, stChamp : string;
begin
  if T.NomTable<>'' then
  begin
    if T.NomTable = 'ECRGUI' then stWhere := ' EG_TYPE="'+W_W+'"'
    else stWhere := ' AG_TYPE="'+W_W+'"';
    Q :=OpenSQL('SELECT * FROM '+T.NomTable+' WHERE '+stWhere,False);
    Q.Insert;
    InitNew(Q);
    for i := 1 to T.NbChamps  do
    begin
      stChamp := T.GetNomChamp(i);
      Q.FindField(stChamp).AsVariant := T.GetValeur(i);
    end;
    Q.Post;
    Ferme (Q);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/05/2004
Modifié le ... : 16/12/2004
Description .. :
Mots clefs ... :
*****************************************************************}
//procedure FusionneFiltre(vNumStd : integer; vTableDes, vTableSrc, vCle, vCondition : string; vEcran : TForm);
procedure TAligneStd.FusionneFiltre(vNumStd : integer; vTableDes, vTableSrc, vCle, vCondition : string);
//{$IFNDEF EAGLSERVER}
var lTraitementFiltre : TTraitementFiltre;
//{$ENDIF}
begin
//{$IFNDEF EAGLSERVER}
  lTraitementFiltre := TTraitementFiltre.Create;
  lTraitementFiltre.FStTableSrc := vTableSrc;
  lTraitementFiltre.FStTableDes := vTableDes;
  lTraitementFiltre.FNumStd     := vNumStd;
  lTraitementFiltre.Execute;
  FreeAndNil(lTraitementFiltre);
//{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TAligneStd.CreationGeneraux(vInNumPlan: integer; vStNumCompte: string): Boolean;
var i : integer;
        lSuffixe       : string;
        lNomChamp      : string;
        lValChamp      : Variant;
        TobGenerauxRef : Tob;
        TobGeneraux    : Tob;
        lQuery         : TQuery;
        lSt            : string;
begin
  Result := False;
  try
    TobGenerauxRef := Tob.Create('GENERAUXREF', nil, -1);
    TobGeneraux    := Tob.Create('GENERAUX', nil, -1);
    lSt := 'SELECT * FROM GENERAUXREF WHERE ' +
           'GER_NUMPLAN = ' + IntToStr( vInNumPlan ) + ' AND ' +
           'GER_GENERAL = "' + BourreLess( vStNumCompte, fbGene) + '"';

    lQuery := OpenSql( lSt, True);
    if not lQuery.Eof then
    begin
      if TobGenerauxRef.SelectDB( '', lQuery, false) then
      begin
        for i := 1 to TobGenerauxRef.NbChamps do
        begin
          lNomChamp := TobGenerauxRef.GetNomChamp(i);
          lSuffixe  := ExtractSuffixe( TobGenerauxRef.GetNomChamp(i) );
          lValChamp := TobGenerauxRef.GetValeur(i);
          if TobGeneraux.FieldExists( 'G_' + lSuffixe ) then
          begin
            if lSuffixe = 'GENERAL' then lValChamp := BourreEtLess ( lValChamp, fbGene);
            TobGeneraux.PutValue('G_' + lSuffixe, lValChamp );
          end;
        end;
        TobGeneraux.InsertDb( nil );
        Result := True;
      end;
    end;

  finally
    Ferme( lQuery );
    FreeAndNil( TobGenerauxRef );
    freeAndNil( TobGeneraux );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/05/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TAligneStd.ChargeAligneStdAvecTob( vTobAligneStd: Tob);
begin
  AjouteTAF( alGeneraux,    vTobAligneStd.Detail[0].GetValue('NUMSTD'));
  AjouteTAF( alTiers,       vTobAligneStd.Detail[1].GetValue('NUMSTD'));
  AjouteTAF( alJournaux,    vTobAligneStd.Detail[2].GetValue('NUMSTD'));
  AjouteTAF( alGuides,      vTobAligneStd.Detail[3].GetValue('NUMSTD'));
  AjouteTAF( alSections,    vTobAligneStd.Detail[4].GetValue('NUMSTD'));
  AjouteTAF( alStdEdition,  vTobAligneStd.Detail[5].GetValue('NUMSTD'));
  AjouteTAF( alRuptEdition, vTobAligneStd.Detail[6].GetValue('NUMSTD'));
  AjouteTAF( alCpteCorresp, vTobAligneStd.Detail[7].GetValue('NUMSTD'));
  AjouteTAF( alTabLibres,   vTobAligneStd.Detail[8].GetValue('NUMSTD'));
  AjouteTAF( alTva,         vTobAligneStd.Detail[9].GetValue('NUMSTD'));
  AjouteTAF( alLibelleAuto, vTobAligneStd.Detail[10].GetValue('NUMSTD'));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/05/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TAligneStd.FOnProgressFormInformation( Sender : TObject ; Erreur : integer ; Msg : string );
begin
  MoveCurProgressForm( Msg );
end;

////////////////////////////////////////////////////////////////////////////////

end.
