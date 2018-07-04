{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/11/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPSTATUTDOSSIER ()
Mots clefs ... : TOF;CPSTATUTDOSSIER
*****************************************************************}
Unit CPSTATUTDOSSIER_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
{$IFDEF EAGLCLIENT}
     eMul,
     MaineAGL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_Main,
     PrintDBG, // PrintDBG
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Vierge,
     extctrls,    // TImage
     Windows,     // TRect
     Grids,       // TGridDrawState
     Graphics,    // TCanvas
     AglInit,     // TheData
     ParamSoc,    // GetParamSocSecur
     ed_Tools,    // MoveCurProgressForm
     Ent1,        // VH^.
     uTobDebug,   // TobDebug
     ZCumul,      // FZCumul
     HSysMenu,    // THSystemMenu
     Htb97,       // TToolBarButton97
     uLibWindows, // IIF

     uTob,
     UTOF ;

Type

  TCPStatutLiasse = class (TObject)

  private
    FTobGene      : Tob; // Liste des comptes généraux
    FTobCollectif : Tob; // Liste des comptes collectifs
    FTobRub       : Tob; //
    FTobEcr       : Tob; //

    FLastError     : integer;
    FAvecAffichage : Boolean;

    FStCptBilanOuv : string; // Compte de bilan d'ouverture
    FStCptBilanFer : string; // Compte de bilan de fermeture

    FZCumul : TZCumul;

    function  EstVraiErreur4(vStGeneral, vStRubrique : string): Boolean;
    procedure TraiteInfoCompte(vTobRub : Tob; vSt: string);
    procedure TraiteGenerauxSansAgregat;
    //procedure TraiteCompteCollectif;
    procedure GenereErreurCollectif(vTobErreur, vTobCollectif : Tob);

  public
    FTobErreur : Tob;

    constructor Create;
    procedure   Execute;
    destructor  Destroy; override;

    property LastError : integer read FLastError;
    property AvecAffichage : Boolean read FAvecAffichage write FAvecAffichage;
  end;


  TOF_CPSTATUTDOSSIER = Class (TOF)

    FListe : THGrid;
    CBOk1  : TCheckBox;
    CBOk2  : TCheckBox;
    CBOk3  : TCheckBox;
    CBOk4  : TCheckBox;
    IM1    : TImage;
    IM2    : TImage;

    BCherche  : TToolBarButton97;
    BImprimer : TToolBarButton97;

    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnPostDrawCellFListe (ACol, ARow : Integer; Canvas : TCanvas; AState : TGridDrawState);
    procedure OnClickCB       ( Sender : TObject );
    procedure OnClickBCherche ( Sender : TObject );
    procedure OnClickBImprimer( Sender : TObject );
    procedure OnDblClickFListe( Sender : TObject );
    procedure OnKeyDownEcran  ( Sender : TObject; var Key : Word; Shift : TShiftState); virtual;
    procedure OnKeyDownFListe ( Sender : TObject; var Key : Word; Shift : TShiftState);

  private
    FTobFListe : Tob;
    FStTitreEcran : string;

    procedure CalculStatutLiasse;
    procedure AfficheErreur;
    procedure StatutEcran( vBoARecalculer : boolean );



  end ;

procedure ControleStatutLiasse;
procedure SelectAgregatDictionnaire( vStSelect : string; vStWhere : string; vDateValidite : TDateTime; vStDictionnaire : string ; vTobRub : Tob);

Implementation

uses uTofConsEcr;

const cColErreur    = 1;
      cColCompte    = 2;
      cColLibelle   = 3;
      cColAgregat   = 4;
      cColComptable = 5;
      cColFiscal    = 6;
      cColMvt       = 7;
      cColSolder    = 8;

////////////////////////////////////////////////////////////////////////////////
procedure ControleStatutLiasse;
begin
  V_Pgi.ZoomOle := True;
  AGLLanceFiche('CP', 'CPSTATUTDOSSIER', '', '', '');
  V_Pgi.ZoomOle := False;
end;

////////////////////////////////////////////////////////////////////////////////
procedure SelectAgregatDictionnaire( vStSelect : string; vStWhere : string; vDateValidite : TDateTime; vStDictionnaire : string ; vTobRub : Tob);
var lTobTemp : Tob;
    i : integer;
    lStOldRub : string;
begin
  vTobRub.ClearDetail;
  vTobRub.LoadDetailFromSQL('SELECT ' + vStSelect + ' FROM RUBRIQUE WHERE ' +
                            IIF( vStDictionnaire <> '', 'RB_FAMILLES LIKE "%' + vStDictionnaire + '%" AND ', '') +
                            'RB_NATRUB = "CPT" AND ' +
                            'RB_CLASSERUB = "CPT" AND ' +
                            'RB_TYPERUB = "GEN" AND ' +
                            '(RB_PREDEFINI <> "DOS" OR ' +
                            '(RB_PREDEFINI = "DOS" AND RB_NODOSSIER = "'+V_PGI.NoDossier+'")) AND ' +
                            'RB_DATEVALIDITE >= "' + UsDateTime(vDateValidite) + '" AND ' +
                            vStWhere + ' ORDER BY RB_RUBRIQUE, RB_PREDEFINI ASC, RB_DATEVALIDITE ASC');

  lStOldRub := '';
  for i := vTobRub.Detail.Count-1 downto 0 do
  begin
    lTobTemp := vTobRub.Detail[i];
    if lTobTemp.GetString('RB_RUBRIQUE') = lStOldRub then
      lTobTemp.Free
    else
      lStOldRub := lTobTemp.GetString('RB_RUBRIQUE');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/10/2006
Modifié le ... : 08/11/2006
Description .. : Controle l'état du dossier par rapport à sa liasse fiscale
Suite ........ : Renvoi 0 si Pas de problème
Suite ........ :        1 si Comptes non mvts pas dans une fourchette
Suite ........ :        2 si Comptes mvts mais soldés pas dans une fourchette
Suite ........ :        3 si Comptes non soldés pas dans une fouchette
Suite ........ :        4 si Comptes non soldés dans un sens différent de son agrégat
Suite ........ :        5 Statut à recalculer
Mots clefs ... :
*****************************************************************}
{ TCPStatutLiasse }
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/11/2006
Modifié le ... : 08/11/2006
Description .. :
Mots clefs ... :
*****************************************************************}
constructor TCPStatutLiasse.Create;
begin
  FTobErreur    := Tob.Create('ERREUR', nil, -1);
  FTobGene      := Tob.Create('GENE', nil, -1);
  FTobCollectif := Tob.Create('COLLECTIF', nil, -1);
  FTobRub       := Tob.Create('RUB', nil, -1);
  FTobEcr       := Tob.Create('ECR', nil, -1);

  FZCumul := TZCumul.Create;
  FZCumul.ModeFonc    := mfMemory;
  FZCumul.AvecLibelle := True;
  FZCumul.InitCriteres('N', '' , '' , '(' + DateToStr(VH^.EnCours.Deb) + ')(' + DateToStr(VH^.EnCours.Fin) + ')', '');

  FLastError  := 0;
  FAvecAffichage := False;

  FStCptBilanOuv := GetParamSocSecur('SO_OUVREBIL', '');
  FStCptBilanFer := GetParamSocSecur('SO_FERMEBIL', '');

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/11/2006
Modifié le ... : 08/11/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TCPStatutLiasse.Execute;
var lesCptes : TStringList;
    lTobTemp : Tob;
    i,j : integer;
    lStRub  : string;
    lStLiasseDossier : string;
    lStDateDeb       : string;
    lStDateFin       : string;
begin
  lStLiasseDossier := GetParamSocSecur('SO_CPCONTROLELIASSE', '');
  if lStLiasseDossier = '' then
  begin
    PgiInfo('Vous devez renseigner la liasse fiscale utilisée par le dossier.','Paramétres société');
    Exit;
  end;

  lStDateDeb := DateToStr(VH^.EnCours.Deb);
  lStDateFin := DateToStr(VH^.EnCours.Fin);

  LesCptes   := TStringList.Create;

  try
    FTobGene.ClearDetail;
    FTobGene.LoadDetailFromSQL('SELECT G_GENERAL, G_LIBELLE, G_TOTDEBE, G_TOTCREE, G_COLLECTIF FROM GENERAUX WHERE ' +
                               'G_GENERAL <> "' + FStCptBilanOuv + '" AND ' +
                               'G_GENERAL <> "' + FStCptBilanFer + '" ORDER BY G_GENERAL');

    FTobCollectif.ClearDetail;
    FTobCollectif.LoadDetailFromSQL('SELECT G_GENERAL, G_LIBELLE, ' +
                                    '0 TOTSD, "-" RUBSD, "" NOMRUBSD, ' +
                                    '0 TOTSC, "-" RUBSC, "" NOMRUBSC, ' +
                                    //'0 TOTSM, "" RUBSM, "" NOMRUBSM ' +
                                    'G_TOTDEBE, G_TOTCREE ' +
                                    'FROM GENERAUX WHERE ' +
                                    'G_COLLECTIF = "X" AND ' +
                                    'G_GENERAL <> "' + FStCptBilanOuv + '" AND ' +
                                    'G_GENERAL <> "' + FStCptBilanFer + '" ORDER BY G_GENERAL');

    FTobEcr.ClearDetail;
    FTobEcr.LoadDetailFromSQL('SELECT G_GENERAL FROM GENERAUX WHERE ((EXISTS(SELECT E_GENERAL FROM ECRITURE WHERE ' +
            'E_GENERAL = G_GENERAL)) OR (EXISTS(SELECT Y_GENERAL FROM ANALYTIQ WHERE Y_GENERAL = G_GENERAL)))');


    SelectAgregatDictionnaire('RB_RUBRIQUE, RB_LIBELLE, RB_COMPTE1, RB_SIGNERUB',
                              '', VH^.EnCours.Fin, lStLiasseDossier, FTobRub);

    InitMoveProgressForm(nil, 'Contrôle du statut du dossier (liasse fiscale)', 'Traitement en cours...', FTobRub.Detail.Count-1, True, True);
    for i := 0 to FTobRub.Detail.Count -1 do
    begin
      MoveCurProgressForm('Veuillez patienter...');
      lTobTemp := FTobRub.Detail[i];
      lStRub := lTobTemp.GetString('RB_RUBRIQUE');

      LesCptes.Clear;
      FZCumul.GetValeur('RUBRIQUE', lStRub, LesCptes);
      if FZCumul.LastError = 0 then
      begin
        // Parcours des informations des comptes
        for j := 0 to LesCptes.Count - 1 do
          TraiteInfoCompte(lTobTemp, LesCptes[j]);
      end
      else
        FLastError := 5;
    end;

    // Traitement de la liste des comptés généraux sans agrégat
    TraiteGenerauxSansAgregat;

    //TraiteCompteCollectif;

    if FTobErreur.Detail.Count > 0 then
      FTobErreur.Detail.Sort('GENERAL');

  finally
    LesCptes.Free;
    FiniMoveProgressForm;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/11/2006
Modifié le ... : 08/11/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TCPStatutLiasse.TraiteInfoCompte(vTobRub : Tob; vSt : string);
var lStInfoCpt : string;
    lStCompte  : string;
    lStLibelle : string;
    lStSuffixe : string;
    lStRubrique: string;

    lDbSD      : Double;
    lDbSC      : Double;
    lDbTD      : Double;
    lDbTC      : Double;
    lDbFiscal  : Double;
    lInCodification : integer;

    lTobTemp   : Tob;
    lTobErreur : Tob;
    lTobCollectif : Tob;

begin
  lStInfoCpt  := vSt;
  lStCompte   := ReadTokenPipe(lStInfoCpt, ':');
  lStRubrique := vTobRub.GetString('RB_RUBRIQUE');

  // Supression du compte dans FTobGene car il est bien affecté à une liasse
  lTobTemp := FTobGene.FindFirst(['G_GENERAL'],[lStCompte], False);
  if lTobTemp <> nil then
    lTobTemp.Free;

  ReadTokenPipe(lStInfoCpt, ':'); // Sens Mixte
  lDBSC  := Valeur( ReadTokenPipe(lStInfoCpt, ':') );           // Sens Créditeur
  lDbSD  := Valeur( ReadTokenPipe(lStInfoCpt, ':') );           // Sens Débiteur
  lDbTC  := Valeur( ReadTokenPipe(lStInfoCpt, ':') );           // Total Crédit du Compte
  lDbTD  := Valeur( ReadTokenPipe(lStInfoCpt, ':') );           // Total Débit du Compte
  lInCodification := StrToInt(ReadTokenPipe(lStInfoCpt, ':'));  // Codification
  lStLibelle := Copy(lStInfoCpt, 5, Length(lStInfoCpt));

  lTobCollectif := FTobCollectif.FindFirst(['G_GENERAL'],[lStCompte], False);
  if lTobCollectif <> nil then
  begin
    lTobCollectif.PutValue('TOTSD', lDBSD );
    lTobCollectif.PutValue('TOTSC', lDBSC );

    if lInCodification = 5 then // SD
    begin
      lTobCollectif.PutValue('RUBSD', 'X');
      lTobCollectif.PutValue('NOMRUBSD', lStRubrique);
    end;

    if lInCodification = 6 then // SC
    begin
      lTobCollectif.PutValue('RUBSC', 'X');
      lTobCollectif.PutValue('NOMRUBSC', lStRubrique);
    end;
  end;

  lStSuffixe := '';
  lDbFiscal  := 0;

  case lInCodification of
    7 : begin // (SM)
          Exit;
        end;

    6 : begin // (SC)
          if (lDbTD - lDBTC) <= 0 then
            Exit
          else
          begin
            lDbFiscal  := lDbSC;
            lStSuffixe := '(SC)';
          end;
        end;

    5 : begin // (SD)
          if (lDbTD - lDBTC) >= 0 then
            Exit
          else
          begin
            lDbFiscal  := lDbSD;
            lStSuffixe := '(SD)';
          end;
        end;

    3 : begin // (TC) non traitée
        end;

    2 : begin // (TD) non traitée
        end;

    0 : Exit;

  else
  end;

  // On arrive ici alors détection eventuelle d'une Erreur 4
  if EstVraiErreur4( lStCompte, lStRubrique) then
  begin
    FLastError := 4;
    lTobErreur := Tob.Create('', FTobErreur, -1);
    lTobErreur.AddChampSupValeur('ERREUR', '4');
    lTobErreur.AddChampSupValeur('GENERAL', lStCompte);
    lTobErreur.AddChampSupValeur('LIBELLE', lStLibelle);
    lTobErreur.AddChampSupValeur('AGREGAT', lStRubrique);

    // GCO - 23/07/2007 - FQ 19243
    if lTobCollectif <> nil then
    begin
      if (lTobCollectif.GetString('RUBSD') = '-') and (lDbSD <> 0) then // Erreur sur SD
        lTobErreur.AddChampSupValeur('COMPTABLE', StrfMontant(lTobCollectif.GetDouble('G_TOTDEBE'), 13, V_Pgi.OkDecV, '', True) + ' D');

      if (lTobCollectif.GetString('RUBSC') = '-') and (lDbSC <> 0) then // Erreur sur SC
        lTobErreur.AddChampSupValeur('COMPTABLE', StrfMontant(lTobCollectif.GetDouble('G_TOTCREE'), 13, V_Pgi.OkDecV, '', True) + ' C');
    end
    else
    begin
      if (lDbTD - lDbTC) >= 0 then
        lTobErreur.AddChampSupValeur('COMPTABLE', StrFMontant(lDbTD - lDbTC, 13, V_Pgi.OkDecV, '', True) + ' D')
      else
        lTobErreur.AddChampSupValeur('COMPTABLE', StrFMontant(Abs(lDbTD - lDbTC), 13, V_Pgi.OkDecV, '', True) + ' C');
    end;

    lTobErreur.AddChampSupValeur('FISCAL', StrfMontant(lDbFiscal, 13, V_Pgi.OkDecV, '', True) + ' ' + lStSuffixe);

    if (Arrondi(lDbTD,2) <> 0) or (Arrondi(lDbTC,2) <> 0) then
      lTobErreur.AddChampSupValeur('MVT', 'X')
    else
      lTobErreur.AddChampSupValeur('MVT', '-');

    if Arrondi(lDbTD - lDbTC, 2) = 0 then
      lTobErreur.AddChampSupValeur('SOLDER', 'X')
    else
      lTobErreur.AddChampSupValeur('SOLDER', '-');
  end;    
end;
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/11/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TCPStatutLiasse.EstVraiErreur4(vStGeneral, vStRubrique : string): Boolean;
var i,j : integer;
    lTobTemp : Tob;
    lStTemp     : string;
    lStRubrique : string;
    lStGeneral  : string;
    LesCptes : TStringList;
    lDbTD    : Double;
    lDbTC    : Double;
    lInCodification : integer;
begin
  LesCptes := TStringList.Create;
  try
    Result := True;
    for i := 0 to FTobRub.Detail.Count -1 do
    begin
      lTobTemp := FTobRub.Detail[i];
      lStRubrique := lTobTemp.GetString('RB_RUBRIQUE');
      if vStRubrique <> lStRubrique then
      begin
        LesCptes.Clear;
        FZCumul.GetValeur('RUBRIQUE', lStRubrique, LesCptes);
        for j := 0 to LesCptes.Count - 1 do
        begin
          lStTemp := LesCptes[j];
          lStGeneral := ReadTokenPipe(lStTemp, ':');
          if lStGeneral <> vStGeneral then Continue;

          // Même Compte général avec Rubrique différente
          ReadTokenPipe(lStTemp, ':');                              // Sens Mixte
          ReadTokenPipe(lStTemp, ':');                              // Sens Créditeur
          ReadTokenPipe(lStTemp, ':');                              // Sens Débiteur
          lDbTC := Valeur( ReadTokenPipe(lStTemp, ':') );           // Total Crédit du Compte
          lDbTD := Valeur( ReadTokenPipe(lStTemp, ':') );           // Total Débit du Compte
          lInCodification := StrToInt(ReadTokenPipe(lStTemp, ':')); // Codification

          case lInCodification of
            7 : begin // (SM)
                  Result := False;
                  Exit;
                end;

            6 : begin // (SC)
                  if (lDbTD - lDBTC) <= 0 then
                  begin
                    Result := False;
                    Exit;
                  end;

                end;

            5 : begin // (SD)
                  if (lDbTD - lDBTC) >= 0 then
                  begin
                    Result := False;
                    Exit;
                  end;
                end;

            3 : begin end; // ??

            2 : begin end; // ??

            else
          end;
        end;  
      end;
    end;

  finally
    LesCptes.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/11/2006
Modifié le ... : 08/11/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TCPStatutLiasse.TraiteGenerauxSansAgregat;
var lTobErreur : Tob;
    i : integer;
    lInErreur : integer;
    lSolde : Double;
    lTobCollectif : Tob;
    lTobGene : Tob;
    lStGene : string;
begin
  if FTobGene.Detail.Count = 0 then Exit;
  for i := 0 to FTobGene.Detail.Count -1 do
  begin
    lTobGene := FTobGene.Detail[i];
    lStGene  := lTobGene.GetString('G_GENERAL');

    // GCO - 23/07/207 - FQ 19243
    lTobCollectif := FTobCollectif.FindFirst(['G_GENERAL'],[lStGene], False);

    lTobErreur := Tob.Create('', FTobErreur, -1);
    lTobErreur.AddChampSupValeur('ERREUR', '0');
    lTobErreur.AddChampSupValeur('GENERAL', lStGene);
    lTobErreur.AddChampSupValeur('LIBELLE', lTobGene.GetString('G_LIBELLE'));
    lTobErreur.AddChampSupValeur('AGREGAT', '');
    lTobErreur.AddChampSupValeur('COMPTABLE', '');
    lTobErreur.AddChampSupValeur('FISCAL', '');

    if FTobEcr.FindFirst(['G_GENERAL'],[lStGene], False) = nil then
    begin // Non mouvementés hors agrégat
      lTobErreur.AddChampSupValeur('MVT', '-');
      lTobErreur.AddChampSupValeur('SOLDER', 'X');
      lInErreur := 1;
    end
    else
    begin
      lTobErreur.AddChampSupValeur('MVT', 'X');
      lSolde := Arrondi(lTobGene.GetDouble('G_TOTDEBE') - lTobGene.GetDouble('G_TOTCREE'),2);
      if lSolde >= 0 then
        lTobErreur.SetString('COMPTABLE', StrfMontant(lSolde, 13, V_Pgi.OkDecV, '', True) + ' D')
      else
        lTobErreur.SetString('COMPTABLE', StrfMontant(lSolde, 13, V_Pgi.OkDecV, '', True) + ' C');

      if lSolde = 0 then
      begin
       // Mouvementés, soldés hors agrégat
       lTobErreur.AddchampSupValeur('SOLDER', 'X');
       lInErreur := 2;
      end
      else
      begin
        // Mouvementés, non soldés hors agrégat
        lTobErreur.AddchampSupValeur('SOLDER', '-');
        lInErreur := 3;
      end;
    end;

    lTobErreur.SetString('ERREUR', IntToStr(lInErreur));

    if (lTobCollectif <> nil) and (lInErreur = 3) then
      GenereErreurCollectif( lTobErreur, lTobCollectif );

    if lInErreur > FLastError then
      FLastError := lInErreur;

  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/12/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
(*
procedure TCPStatutLiasse.TraiteCompteCollectif;
var i          : integer;
    lTobCollectif : Tob;
    lTobErreur : Tob;
    lInErreur  : Integer;
    lDbSC      : Double;
    lDBSD      : Double;
begin
  for i := 0 to FTobCollectif.Detail.Count -1 do
  begin
    lTobTemp := FTobCollectif.Detail[i];

    lDBSD := Arrondi(lTobTemp.GetDouble('TOTSD'), 2);
    lDBSC := Arrondi(lTobTemp.GetDouble('TOTSC'), 2);

    if ((lTobTemp.GetString('RUBSD') = '-') and (lDbSD <> 0)) or
       ((lTobTemp.GetString('RUBSC') = '-') and (lDbSC <> 0)) then
    begin
      lInErreur := 4;
      lTobErreur := Tob.Create('', FTobErreur, -1);
      lTobErreur.AddChampSupValeur('ERREUR', IntToStr(lInErreur));
      lTobErreur.AddChampSupValeur('GENERAL', lTobTemp.GetString('G_GENERAL'));
      lTobErreur.AddChampSupValeur('LIBELLE', lTobTemp.GetString('G_LIBELLE'));
      lTobErreur.AddChampSupValeur('AGREGAT', '');
      lTobErreur.AddChampSupValeur('COMPTABLE', '');
      lTobErreur.AddChampSupValeur('MVT', 'X');
      lTobErreur.AddChampSupValeur('SOLDER', '-');

      if (lTobTemp.GetString('RUBSD') = '-') and (lDbSD <> 0) then // Erreur sur SD
        lTobErreur.SetString('COMPTABLE', StrfMontant(vTobCollectif.GetDouble('G_TOTDEBE'), 13, V_Pgi.OkDecV, '', True) + ' D')
      else
        lTobErreur.SetString('COMPTABLE', StrfMontant(vTobCollectif.GetDouble('G_TOTCREE'), 13, V_Pgi.OkDecV, '', True) + ' C');

      if lInErreur > FLastError then
        FLastError := 4;
    end;
  end;
end;*)

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/11/2006
Modifié le ... : 08/11/2006
Description .. :
Mots clefs ... :
*****************************************************************}
destructor TCPStatutLiasse.Destroy;
begin
  FreeAndNil( FTobErreur );
  FreeAndNil( FTobGene );
  FreeAndNil( FTobCollectif );
  FreeAndNil( FTobRub );
  FreeAndNil( FTobEcr );
  FreeAndNil( FZCumul );
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/11/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.OnArgument (S : String ) ;
begin
  Inherited ;
  FListe := THGrid(GetControl('FLISTE', True));
  FListe.PostDrawCell := OnPostDrawCellFListe;

  FListe.OnKeyDown  := OnKeyDownFListe;
  FListe.OnDblClick := OnDblClickFListe;

  FListe.ColAligns[cColErreur]    := TaCenter;
  FListe.ColAligns[cColCompte]    := TaCenter;
  FListe.ColAligns[cColLibelle]   := TaLeftJustify;
  FListe.ColAligns[cColAgregat]   := TaCenter;
  FListe.ColAligns[cColComptable] := TaRightJustify;
  FListe.ColAligns[cColFiscal]    := TaRightJustify;

  FListe.ColTypes[cColMvt]        := 'B';
  FListe.ColAligns[cColMvt]       := TaCenter;
  FListe.ColFormats[cColMvt]      := IntToStr( Byte( csCheckBox ) );

  FListe.ColTypes[cColSolder]     := 'B';
  FListe.ColAligns[cColSolder]    := TaCenter;
  FListe.ColFormats[cColSolder]   := IntToStr( Byte( csCheckBox ) );

  CBOk1 := TCheckBox(GetControl('CBOK1', True));
  CBOk2 := TCheckBox(GetControl('CBOK2', True));
  CBOk3 := TCheckBox(GetControl('CBOK3', True));
  CBOk4 := TCheckBox(GetControl('CBOK4', True));

  CBOk1.OnClick := OnClickCB;
  CBOk2.OnClick := OnClickCB;
  CBOk3.OnClick := OnClickCB;
  CBOk4.OnClick := OnClickCB;

  IM1 := TImage(GetControl('IM1', True));
  IM2 := TImage(GetControl('IM2', True));

  BCherche  := TToolBarButton97(Getcontrol('BCHERCHE', True));
  BCherche.OnClick  := OnClickBCherche;

  BImprimer := TToolBarButton97(Getcontrol('BIMPRIMER', True));
  BImprimer.OnClick := OnClickBImprimer;

{$IFDEF EAGLCLIENT}
  BImprimer.Visible := False;
{$ENDIF}

  FTobFListe := Tob.Create('', nil, -1);

  FStTitreEcran := TraduireMemoire('Statut du dossier (liasse fiscale)') + ' - ' +
                   RechDom('TTRUBFAMILLE', GetParamSocSecur('SO_CPCONTROLELIASSE', ''), false);

  Ecran.OnKeyDown := OnKeyDownEcran;

  StatutEcran( False );

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/11/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/11/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.StatutEcran(vBoARecalculer: boolean);
begin
  Ecran.Caption := FStTitreEcran;
  if vBoARecalculer then
  begin
    Ecran.Caption := Ecran.Caption + ' (à recalculer)';
    BCherche.Glyph := IM2.Picture.BitMap;
  end
  else
    BCherche.Glyph := IM1.Picture.BitMap;

  UpdateCaption(Ecran);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/11/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.OnPostDrawCellFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var TheRect : TRect;
begin
  if (ACol = cColAgregat) or
     (ACol = cColComptable) or
     (ACol = cColFiscal) then
  begin
    if ((Trim(FListe.Cells[ACol, ARow]) = '') or
        (Valeur(FListe.Cells[ACol, ARow]) = 0)) then
    begin
      TheRect := FListe.CellRect(ACol, ARow);
      Canvas.TextRect(TheRect, TheRect.Left, TheRect.Top, '');
      Canvas.Brush.Color := FListe.FixedColor;
      Canvas.Brush.Style := bsBDiagonal;
      Canvas.Pen.Color := FListe.FixedColor;
      Canvas.Pen.Mode := pmCopy;
      Canvas.Pen.Style := psClear;
      Canvas.Pen.Width := 1;
      Canvas.Rectangle(TheRect);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/11/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.OnClickCB(Sender: TObject);
begin
  AfficheErreur;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/11/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.OnLoad ;
begin
  BCherche.Click;
end ;



////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/11/2006
Modifié le ... : 05/12/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.OnClickBCherche(Sender: TObject);
begin
  CalculStatutLiasse;
  StatutEcran(False);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/11/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.OnClickBImprimer(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
  PrintDBGrid (FListe, nil, Ecran.Caption, '');
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/11/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of

    VK_F5 : if FListe.Focused then FListe.OnDblClick(Self);

    VK_F9 : BCherche.Click;

    80 :  if Shift = [ssCtrl] then // Ctrl + P
            BImprimer.Click;
  else
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/11/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.OnKeyDownFListe(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if FListe.SortEnabled then
  begin
    if (Shift = [ssAlt]) and (Key in [Ord('1')..Ord('9')]) then
    begin
      case key of
        Ord('1') : FListe.SortGrid(1, (FListe.SortedCol = 1) xor FListe.SortDesc);
        Ord('2') : FListe.SortGrid(2, (FListe.SortedCol = 2) xor FListe.SortDesc);
        Ord('3') : FListe.SortGrid(3, (FListe.SortedCol = 3) xor FListe.SortDesc);
        Ord('4') : FListe.SortGrid(4, (FListe.SortedCol = 4) xor FListe.SortDesc);
        Ord('5') : FListe.SortGrid(5, (FListe.SortedCol = 5) xor FListe.SortDesc);
        Ord('6') : FListe.SortGrid(6, (FListe.SortedCol = 6) xor FListe.SortDesc);
        Ord('7') : FListe.SortGrid(7, (FListe.SortedCol = 7) xor FListe.SortDesc);
        Ord('8') : FListe.SortGrid(8, (FListe.SortedCol = 8) xor FListe.SortDesc);
        Ord('9') : FListe.SortGrid(9, (FListe.SortedCol = 9) xor FListe.SortDesc);
      else
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/11/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.CalculStatutLiasse;
var lCPStatutLiasse : TCPStatutLiasse;
    lTobTemp : Tob;
    i : integer;
begin
  lCPStatutLiasse := TCPStatutLiasse.Create;
  try
    lCPStatutLiasse.Execute;
    case lCPStatutLiasse.FLastError of
      0   : VH^.BStatutLiasse.ImageIndex := 0; // Vert
      1,2 : VH^.BStatutLiasse.ImageIndex := 1; // Orange
      3,4 : VH^.BStatutLiasse.ImageIndex := 2; // Rouge
      5   : VH^.BStatutLiasse.ImageIndex := 3; // Bleu
    else
    end;
    // Mise à jour du paramsoc qui conserve la couleur du Statut
    SetParamSoc('SO_CPSTATUTDOSSIERLIASSE', VH^.BStatutLiasse.ImageIndex);
    
    FTobFListe.ClearDetail;
    if lCPStatutLiasse.FTobErreur.Detail.Count > 0 then
    begin
      for i := 0 to lCPStatutLiasse.FTobErreur.Detail.Count - 1 do
      begin
        lTobTemp := Tob.Create('', FTobFListe, -1);
        lTobTemp.Dupliquer(lCPStatutLiasse.FTobErreur.Detail[i], True, True);
      end;

      AfficheErreur;
    end;
  finally
    Ecran.Caption := FStTitreEcran;
    UpdateCaption(Ecran);
    lCPStatutLiasse.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/11/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.AfficheErreur;
var i : integer;
    lBoOk1 : Boolean;
    lBoOk2 : Boolean;
    lBoOk3 : Boolean;
    lBoOk4 : Boolean;
    lInErreur : integer;
begin
  if FTobFListe = nil then Exit;

  try
    lBoOk1 := CBOk1.Checked;
    lBoOk2 := CBOk2.Checked;
    lBoOk3 := CBOk3.Checked;
    lBoOk4 := CBOk4.Checked;

    FListe.VidePile(False);
    FListe.BeginUpdate;
    FListe.RowCount := 2;
    for i := 0 to FTobFListe.Detail.count -1 do
    begin
      lInErreur := FTobFListe.Detail[i].GetInteger('ERREUR');

      if (lBoOk1 and (lInErreur = 1)) or
         (lBoOk2 and (lInErreur = 2)) or
         (lBoOk3 and (lInErreur = 3)) or
         (lBoOk4 and (lInErreur = 4)) then
      begin
        FListe.InsertRow(FListe.RowCount-1);
        FListe.Cells[cColErreur, FListe.RowCount-2]    := FTobFListe.Detail[i].Getstring('ERREUR');
        FListe.Cells[cColCompte, FListe.RowCount-2]    := FTobFListe.Detail[i].Getstring('GENERAL');
        FListe.Cells[cColLibelle, FListe.RowCount-2]   := FTobFListe.Detail[i].Getstring('LIBELLE');
        FListe.Cells[cColAgregat, FListe.RowCount-2]   := FTobFListe.Detail[i].Getstring('AGREGAT');
        FListe.Cells[cColComptable, FListe.RowCount-2] := FTobFListe.Detail[i].Getstring('COMPTABLE');
        FListe.Cells[cColFiscal, FListe.RowCount-2]    := FTobFListe.Detail[i].Getstring('FISCAL');
        FListe.Cells[cColMvt, FListe.RowCount-2]       := FTobFListe.Detail[i].Getstring('MVT');
        FListe.Cells[cColSolder, FListe.RowCount-2]    := FTobFListe.Detail[i].Getstring('SOLDER');
      end;
    end;

    if FListe.RowCount > 2 then
      FListe.RowCount := FListe.RowCount-1;

  finally
    FListe.Row := 1;
    THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
    FListe.EndUpdate ;
    if FListe.CanFocus then FListe.SetFocus;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/11/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.OnDblClickFListe(Sender: TObject);
begin
  if FListe.Cells[cColCompte, FListe.Row] <> '' then
  begin
    OperationsSurComptes(FListe.Cells[cColCompte, FListe.Row], VH^.Encours.Code, '','', False);
    StatutEcran( True);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/11/2006
Modifié le ... : 08/11/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTDOSSIER.OnClose ;
begin
  FreeAndNil(FTobFliste);
  Inherited;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/07/2007
Modifié le ... :   /  /    
Description .. : FQ 19243
Suite ........ : On Eclate la ligne d'erreur en deux pour faire une erreur au
Suite ........ : Débit et une au Crédit 
Mots clefs ... :
*****************************************************************}
procedure TCPStatutLiasse.GenereErreurCollectif(vTobErreur, vTobCollectif : Tob);
var lTobErreur : Tob;
begin
  lTobErreur := Tob.Create('', FTobErreur, -1);
  lTobErreur.Dupliquer(vTobErreur, False, True);

  // Info débiteur du Collectif
  vTobErreur.PutValue('COMPTABLE', StrFMontant( vTobCollectif.GetDouble('G_TOTDEBE'), 13, V_Pgi.OkDecV, '', True) + ' D');
  vTobErreur.PutValue('AGREGAT', vTobCollectif.GetValue('NOMRUBSD'));
  vTobErreur.PutValue('FISCAL', StrfMontant(vTobCollectif.GetDouble('TOTSD'), 13, V_Pgi.OkDecV, '', True) + ' ' + '(SD)');

  // Info Créditeur du collectif
  lTobErreur.PutValue('COMPTABLE', StrFMontant( vTobCollectif.GetDouble('G_TOTCREE'), 13, V_Pgi.OkDecV, '', True) + ' C');
  lTobErreur.PutValue('AGREGAT', vTobCollectif.GetValue('NOMRUBSC'));
  lTobErreur.PutValue('FISCAL',  StrfMontant(vTobCollectif.GetDouble('TOTSC'), 13, V_Pgi.OkDecV, '', True) + ' ' + '(SC)');
end;

Initialization
  registerclasses ( [ TOF_CPSTATUTDOSSIER ] ) ;
end.