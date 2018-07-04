{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 24/01/2001
Modifié le ... : 22/04/2003
Description .. : Source TOF de la TABLE : CLOTUREDEF ()
Suite ........ : - 12/03/2001 - CA - Clôture sans AN sans création d'exercice suivant
Suite ........ : - 22/04/2003 - CA - FQ 10410 - vérification absence exercice avant création.
Suite ........ : - 22/04/2003 - CA - FQ 10597 - Libellé exercice.
Suite ........ : - 22/04/2003 - CA - FQ 12259 - Contrôle durée d'exercice
Suite ........ : - 26/05/2003 - CA - FQ 12295 - Mise à jour de S0_CPEXOREF
Suite ........ : - 26/08/2003 - CA - FQ 12644 - Initialisation des champs exercice
Suite ........ : - 12/09/2003 - CA - FQ 12658 - Rechargement des paramètres
Mots clefs ... : TOF;CLOTUREDEF;CLOTURE;À-NOUVEAUX
*****************************************************************}

unit UTofClotureDef;

interface

uses StdCtrls, Controls, Classes, db, forms, sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOF, Ent1, OuvreExo, HTB97, Cloture, Audit, SoldeCpt,
  ParamSoc, uTOB
  {$IFDEF EAGLCLIENT}
  , MaineAGL
  {$ELSE}
  , fe_main
  {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
  {$ENDIF}
  {$IFDEF NETEXPERT}
  {$IFDEF COMPTA}
  , UAssistComsx
  {$ENDIF}
   {$ENDIF}
  , Math;

type
  TOF_CLOTUREDEF = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnAuditClick(Sender: TObject);
    procedure OnAvecANouveauClick(Sender: TObject);
    procedure OnSansANouveauClick(Sender: TObject);
    procedure OnClickArchive(Sender: TObject);
  private
    m_bCreationSuivant: boolean;
    m_bOuvrirSuivant: boolean;     
{$IFDEF CERTIFNF}
    SessValidation : Integer;
    ChecksumValidation : String;
{$ENDIF}        
    function ExerciceSuivantOK: boolean;
    procedure ExecuteClotureDefSansAN;
    function OuvertureExerciceSuivant: boolean;
    function GetInfoExerciceSuivant(var DateDeb: TDateTime; var DateFin:
      TDateTime): boolean;
    function AucuneEcriture(Exo: string): boolean;
  end;

procedure CPLanceClotureDefinitive;

implementation

uses
{$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
{$ENDIF MODENT1}
  ulibCloture,    // TTraitementCloture
{$IFDEF AMORTISSEMENT}
  ImoClo ,
{$ENDIF}
{$IFDEF COMPTA}
{$IFDEF CERTIFNF}
  uLibValidation,
{$ENDIF}
{$ENDIF}
  uLibExercice;   // CControleDureeExercice

procedure CPLanceClotureDefinitive;
begin
  AGLLanceFiche('CP','CLOTUREDEF','','','');
end;

procedure TOF_CLOTUREDEF.OnNew;
begin
  inherited;
end;

procedure TOF_CLOTUREDEF.OnDelete;
begin
  inherited;
end;

procedure TOF_CLOTUREDEF.OnUpdate;
begin
  inherited;
end;

procedure TOF_CLOTUREDEF.OnLoad;
var
  An, Mois, Jour: Word;
  ad, md, NbMois: Word;
  SuivantDeb, SuivantFin: TDateTime;
begin
  m_bCreationSuivant := False;
  m_bOuvrirSuivant := False;
  SetControlText('ENCOURSDEB', DateToStr(VH^.Encours.Deb));
  SetControlText('ENCOURSFIN', DateToStr(VH^.Encours.Fin));
  if GetInfoExerciceSuivant(SuivantDeb, SuivantFin) then
  begin
    SetControlText('SUIVANTDEB', DateToStr(SuivantDeb));
    SetControlText('SUIVANTFIN', DateToStr(SuivantFin));
  end
  else
    {  if VH^.Suivant.Code <> '' then
      begin
        SetControlText('SUIVANTDEB',DateToStr(VH^.Suivant.Deb));
        SetControlText('SUIVANTFIN',DateToStr(VH^.Suivant.Fin));
      end else }
  begin
    m_bCreationSuivant := True;
    SuivantDeb := VH^.Encours.Fin + 1;
    NombrePerExo(VH^.EnCours, md, ad, NbMois);
    DecodeDate(SuivantDeb, An, Mois, Jour);
    if Mois + NbMois - 1 <= 12 then
      SuivantFin := EncodeDate(An, Mois + NbMois - 1, 25)
    else
    begin
      if (Mois + NbMois - 1) - 12 > 12 then
        SuivantFin := EncodeDate(An + 1, 12, 25)
      else
        SuivantFin := EncodeDate(An + 1, (Mois + NbMois - 1) - 12, 25);
    end;
    SuivantFin := FinDeMois(SuivantFin);
    SetControlText('SUIVANTDEB', DateToStr(SuivantDeb));
    SetControlText('SUIVANTFIN', DateToStr(SuivantFin));
    SetControlEnabled('SUIVANTFIN', True);
  end;
  inherited;
end;

procedure TOF_CLOTUREDEF.OnArgument(S: string);
var TopAvec, TopSans : integer ;
begin
  inherited;
  // YMO FQ19220&19221
  Ecran.HelpContext := 7751000 ;

  TToolbarButton97(GetControl('BFERME')).Visible := True;

  TToolbarButton97(GetControl('BAUDIT')).OnClick := OnAuditClick;
  TToolbarButton97(GetControl('BAVEC')).OnClick := OnAvecANouveauClick;
  TToolbarButton97(GetControl('BSANS')).OnClick := OnSansANouveauClick;
  
  if not (ctxPCL in V_PGI.PGIContexte) then
  begin
    TToolbarButton97(GetControl('BSANS')).Visible := False;
    TToolbarButton97(GetControl('HELPBTN')).Visible := False;
    TToolbarButton97(GetControl('BFERME')).Left := TToolbarButton97(GetControl('HELPBTN')).Left;
    TopAvec:=TToolbarButton97(GetControl('BAVEC')).Top;
    TopSANS:=TToolbarButton97(GetControl('BSANS')).Top;
    TToolbarButton97(GetControl('BAVEC')).Top :=TopAvec+((TopSans-TopAvec) div 2);
  end;
  TToolbarButton97(GetControl('BAUDIT')).OnClick := OnAuditClick;
{$IFDEF CERTIFNF}
  //Ajout me pour les archivages
  THCheckBox(GetControl('BARCHIVE')).OnClick := OnClickArchive;
{$ENDIF}
end;
procedure TOF_CLOTUREDEF.OnClickArchive(Sender: TObject);
begin
 if not THCheckBox(GetControl('BARCHIVE')).Checked then
 begin
       If PGIAsk ('En référence au BOI 13 L-1-06 N° 12 du 24 Janvier 2006 Paragraphe 95,'+#10#13+
       ' il est préconisé d''archiver l''ensemble des informations dont la conservation est obligatoire.'+#10#13+
       ' Vous avez choisi de ne pas effectuer un archivage automatique de ces données, la responsabilité '+#10#13+
       ' de le réaliser manuellement vous incombe. Confirmez-vous ce choix ?')  <> mrYes then
               THCheckBox(GetControl('BARCHIVE')).Checked := TRUE;
 end;
end;


procedure TOF_CLOTUREDEF.OnClose;
begin
  inherited;
end;

procedure TOF_CLOTUREDEF.ExecuteClotureDefSansAN;
var
  Q: TQuery;         
{$IFDEF CERTIFNF}
    Temp : String;
{$ENDIF}
begin
  // Ouverture de l'exercice suivant le cas échéant
  if ExerciceSuivantOk then
  begin
    if OuvertureExerciceSuivant then
    begin
      // GCO - FQ 13089 - Suppression des ecritures typées "V" en Cloture Définitive
      ExecuteSQL('DELETE FROM ECRITURE WHERE E_EXERCICE = "' + VH^.Encours.Code + '" AND E_QUALIFPIECE = "V"');
      // FIN GCO

      // Ecritures
      {  BVE 27.08.07
      Mise en place d'une nouvelle validation (NF)}
{$IFDEF COMPTA}
{$IFDEF CERTIFNF}
      Temp := ValidationEcriture('E_EXERCICE="' + VH^.Encours.Code + '" ' +
                'AND E_DATECOMPTABLE>="' + USDateTime(VH^.EnCours.Deb) + '" ' +
                'AND E_DATECOMPTABLE<="' + USDateTime(VH^.EnCours.Fin) + '" ' +
                'AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N"');     
      // On traite la variable de retour :
      try
         SessValidation := StrToInt(ReadTokenST(Temp));
         ChecksumValidation := Temp;
      except
         SessValidation := 0;
         ChecksumValidation := '';
      end;
{$ELSE}
      ExecuteSQL('UPDATE ECRITURE SET E_VALIDE="X" WHERE E_EXERCICE="' + VH^.Encours.Code + '" ' +
                'AND E_DATECOMPTABLE>="' + USDateTime(VH^.EnCours.Deb) + '" ' +
                'AND E_DATECOMPTABLE<="' + USDateTime(VH^.EnCours.Fin) + '" ' +
                'AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N"');
{$ENDIF}
{$ELSE}
      ExecuteSQL('UPDATE ECRITURE SET E_VALIDE="X" WHERE E_EXERCICE="' + VH^.Encours.Code + '" ' +
                'AND E_DATECOMPTABLE>="' + USDateTime(VH^.EnCours.Deb) + '" ' +
                'AND E_DATECOMPTABLE<="' + USDateTime(VH^.EnCours.Fin) + '" ' +
                'AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N"');
{$ENDIF}
      // Clôture du dossier
      Q := OpenSQL('SELECT * FROM EXERCICE WHERE EX_EXERCICE="' +
        VH^.Encours.Code + '"', False);
      if not Q.Eof then
      begin
        Q.Edit;
        Q.FindField('EX_ETATCPTA').AsString := 'CDE';
        Q.Post;
      end;
      Ferme(Q);

      { Mise à jour du VISA sur les comptes }
      ExecuteSQL ('UPDATE GENERAUX SET G_VISAREVISION="-"');

      SetParamSoc('SO_EXOV8', GetColonneSQL('EXERCICE', 'EX_EXERCICE',
        'EX_ETATCPTA="OUV" AND EX_DATEDEBUT="' + USDateTime(VH^.Encours.Fin + 1) +
        '"'));
      { FQ 12294 - CA - 26/05/2003 - Mise à jour de CPEXOREF }
      SetParamSoc('SO_CPEXOREF', VH^.Suivant.Code);
      // ChargeMagExo(False);
      AvertirCacheServer( 'PARAMSOC' ) ;
      AvertirCacheServer( 'EXERCICE' ) ;
      CHARGESOCIETEHALLEY;
      AvertirMultiTable('TTEXERCICE');
      MajTotTousComptes(True, '');
    end;
  end;
end;

function TOF_CLOTUREDEF.ExerciceSuivantOK: boolean;
var
  ad, md, NbMois: Word;
  DateFin: TDateTime;
  ExoSuivant: TExoDate;
begin
  if m_bCreationSuivant then
  begin
    // On vérifie les paramètres saisis par l'utilisateur
    ExoSuivant.Deb := StrToDate(GetControlText('SUIVANTDEB'));
    ExoSuivant.Fin := StrToDate(GetControlText('SUIVANTFIN'));
    NombrePerExo(ExoSuivant, md, ad, NbMois);
    DateFin := ExoSuivant.Fin;
    // CA - 22/04/2003 - FQ 12259
    if not CControleDureeExercice(ExoSuivant.Deb, ExoSuivant.Fin) then
    begin
      Result := False;
    end
    else if (DateFin <> ExoSuivant.Fin) then
    begin
      SetControlText('SUIVANTFIN', DateToStr(DateFin));
      Result := False;
    end else if (ExoSuivant.Fin < ExoSuivant.Deb ) then
    begin
      Result := False;
    end
    else
      Result := True;
    // Sinon, l'exercice a déjà été créé, c'est donc OK
  end
  else
    result := True;
end;

procedure TOF_CLOTUREDEF.OnAuditClick(Sender: TObject);
begin
  LanceAuditPourCloture;
end;                                                         

procedure TOF_CLOTUREDEF.OnAvecANouveauClick(Sender: TObject);
var
  bErreur: boolean;
  vAnnulClo : integer ;
  ExoLog : TExoDate;
begin
 
  if AucuneEcriture(VH^.Encours.Code) then
    exit;
  bErreur := false;
  SetControlEnabled('BAUDIT', False);
  SetControlEnabled('BSANS', False);
  SetControlEnabled('BAVEC', False);
  if ExerciceSuivantOk then
  begin
    if OuvertureExerciceSuivant then
    begin
      if ClotureComptable(True) then
      begin
        // ChargeMagExo(False); { ne charge pas toutes les infos susceptibles d'être modifiées ( notamment exoV8 }
        CHARGESOCIETEHALLEY;
        AvertirMultiTable('TTEXERCICE');

        //09/01/2006 YMO Norme NF
        QuelDateDeExo(VH^.Precedent.Code, ExoLog);
        { BVE 29.08.07 : Mise en place d'un nouveau tracage }
{$IFNDEF CERTIFNF}
        CPEnregistreLog('SOLDE COMPTES GESTION '+DateToStr(ExoLog.Fin));
{$ELSE}
        CPEnregistreJalEvent('CSG','Solde comptes gestion','SOLDE COMPTES GESTION '+DateToStr(ExoLog.Fin));
{$ENDIF}
        QuelDateDeExo(VH^.EnCours.Code, ExoLog);
        { BVE 29.08.07 : Mise en place d'un nouveau tracage }
{$IFNDEF CERTIFNF}
        CPEnregistreLog('REPORT A NOUVEAU '+DateToStr(ExoLog.Deb));
{$ELSE}
        CPEnregistreJalEvent('CRA','Report à nouveau','REPORT A NOUVEAU '+DateToStr(ExoLog.Deb));
{$ENDIF}

        // GCO - 18/09/2006 - Traçage des évènements
        // L'exercice cloturé est devenu l'exercice précédent.
        { BVE 29.08.07 : Mise en place d'un nouveau tracage }
{$IFNDEF CERTIFNF}
        CPEnregistreLog('CLOEXO ' + VH^.Precedent.Code);
{$ELSE}
        CPEnregistreJalEvent('CEX','Clôture exercice','CLOEXO ' + VH^.Precedent.Code);
{$ENDIF}

        //06/12/2006 YMO Norme NF 203 : interdiction d'annuler en cascade
        vAnnulClo:=GetParamSocSecur('SO_CPANNULCLO', False);
        SetParamSoc('SO_CPANNULCLO',Max(vAnnulClo-1,0));
        {$IFDEF NETEXPERT}
        {$IFDEF COMPTA}
        {$IFDEF CERTIFNF}
        // si archivage coché
        if THCheckBox(GetControl('BARCHIVE')).Checked then
                EnvoiExportParDate (DateToStr(VH^.Precedent.Deb), DateToStr(VH^.Precedent.Fin), VH^.Precedent.Code,  'CLO');
        {$ENDIF}
        {$ENDIF}
        {$ENDIF}
      end
      else
        bErreur := True;
    end
    else
    begin
      MessageAlerte('Ouverture de l''exercice suivant impossible.');
      bErreur := True;
    end;
  end
  else
    bErreur := True;
  SetControlEnabled('BAUDIT', bErreur);
  SetControlEnabled('BSANS', bErreur);
  SetControlEnabled('BAVEC', bErreur);
end;

procedure TOF_CLOTUREDEF.OnSansANouveauClick(Sender: TObject);
var
  bErreur: boolean;
  dMr, vAnnulClo: integer;
  ExoLog : TExoDate;

{$IFDEF AMORTISSEMENT}
  ExoClo : TExoDate;
{$ENDIF}

  // GCO - 216/05/2009 - 13049
  lClotureProcess : TTraitementCloture ;
  errId : integer;
begin
  if AucuneEcriture(VH^.Encours.Code) then exit;

  bErreur := false;
  SetControlEnabled('BAUDIT', False);
  SetControlEnabled('BSANS', False);
  SetControlEnabled('BAVEC', False);
  if ExerciceSuivantOk then
  begin
    if OuvertureExerciceSuivant then
      begin
      m_bCreationSuivant := False;
      dMr := mrYes;

      lClotureProcess := TTraitementCloture.Create( Ecran, '', '', True, False);
      try
        // GCO - 25/01/2006 - FQ 13049
        errID := lClotureProcess.JalExtraComptableOk;
        if (errID <> CLO_PASERREUR) then
        begin
          if PgiAsk(lClotureProcess.LastErrorMsg + '#10#13' +
                    'Voulez-vous abandonner le traitement?', 'Clôture comptable') = MrYes then
                    begin
                      SetControlEnabled('BAUDIT', True);
                      SetControlEnabled('BSANS', True);
                      SetControlEnabled('BAVEC', True);
                      Exit;
                    end;
        end;

        errID := lClotureProcess.GeneExtraComptableOk;
        if (errID <> CLO_PASERREUR) then
        begin
          if PgiAsk(lClotureProcess.LastErrorMsg + '#10#13' +
                    'Voulez-vous abandonner le traitement?', 'Clôture comptable') = MrYes then
                    begin
                      SetControlEnabled('BAUDIT', True);
                      SetControlEnabled('BSANS', True);
                      SetControlEnabled('BAVEC', True);
                      Exit;
                    end;
        end;

        // GCO - 04/12/2006 - FQ 18842
        errID := lClotureProcess.CompteBilanOk;
        if (errID <> CLO_PASERREUR) then
        begin
          if PgiAsk(lClotureProcess.LastErrorMsg + '#10#13' +
                    'Voulez-vous abandonner le traitement?', 'Clôture comptable') = MrYes then
                    begin
                      SetControlEnabled('BAUDIT', True);
                      SetControlEnabled('BSANS', True);
                      SetControlEnabled('BAVEC', True);
                      Exit;
                    end;
        end;

        // GCO - 31/10/2006 - FQ 19020
        errID := lClotureProcess.EcritValidOk ;
        if (errID=CLO_ECRNOVALIDE) and
           (PgiAsk('En référence au BOI 13 L-1-06 N° 12 du 24 janvier 2006 paragraphe 27, nous vous rappelons ' + #13#10 +
                   'que vous devez valider les écritures avant la clôture d''exercice. Certaines écritures ' + #13#10 +
                   'courantes ne sont pas validées. Voulez-vous continuer ?', 'Clôture comptable') = MrNo) then
                   begin
                      SetControlEnabled('BAUDIT', True);
                      SetControlEnabled('BSANS', True);
                      SetControlEnabled('BAVEC', True);
                      Exit;
                    end;
                    // ajout me fiche 19691
         MajValidationJournaux( True );


      finally
        lClotureProcess.Free;
      end;
      // FIN GCO

  {$IFDEF AMORTISSEMENT}
      if AuMoinsUneImmo then
      begin
        QuelDateDeExo(GetParamSoc('SO_EXOCLOIMMO'), ExoClo);
        { si la date de dernière clôture immo est antérieure à l'encours, on lance la clôture des immos }
        if VH^.Encours.Deb > ExoClo.Deb then
        begin
          PGIInfo('Le traitement suivant va proposer la clôture impérative des immobilisations avant la clôture comptable.', ECRAN.Caption);
          AfficheClotureImmo;
          QuelDateDeExo(GetParamSoc('SO_EXOCLOIMMO'), ExoClo);
          { si la clôture des immos n'a pas été faite, on sort }
          if VH^.Encours.Deb > ExoClo.Deb then
            dMr := mrNo
          else
            dMr := mrYes;
        end;
        { Fin clôture des immobilisations dans l'option amortissement }
    //      dMr := PGIAsk('La clôture des immobilisations a-t-elle été faite?',Ecran.Caption);
        if dMr = mrNo then
          bErreur := True;
      end;
  {$ENDIF}
      if (dMr = mrYes) then
      begin
        if Transactions(ExecuteClotureDefSansAN, 3) <> oeOK then
        begin
          MessageAlerte('Problème lors de la clôture du dossier.#10#13Clôture impossible.');
          bErreur := True;
        end
        else
        begin
          ChargeMagExo(False);
          AvertirMultiTable('TTEXERCICE');
          // GCO - 20/07/2005 - FQ 15025
          SetParamsoc('SO_CPDERNEXOCLO',VH^.Precedent.Code);
          SetParamSoc('SO_CPEXOREF',VH^.EnCours.Code);
           {$IFDEF NETEXPERT}
           {$IFDEF COMPTA}
           {$IFDEF CERTIFNF}
          // si archivage coché
            if THCheckBox(GetControl('BARCHIVE')).Checked then
                   EnvoiExportParDate (DateToStr(VH^.Precedent.Deb), DateToStr(VH^.Precedent.Fin), VH^.Precedent.Code, 'CLO');
          {$ENDIF}
          {$ENDIF}
          {$ENDIF}
          // FIN GCO
          PGIInfo('L''opération de clôture du dossier est terminée.', Ecran.Caption);

          //06/12/2006 YMO Norme NF
          QuelDateDeExo(VH^.Precedent.Code, ExoLog);   
          { BVE 29.08.07 : Mise en place d'un nouveau tracage }
{$IFNDEF CERTIFNF}
          CPEnregistreLog('SOLDE COMPTES GESTION '+DateToStr(ExoLog.Fin));
{$ELSE}
          CPEnregistreJalEvent('CSG','Solde comptes gestion','SOLDE COMPTES GESTION '+DateToStr(ExoLog.Fin));
{$ENDIF}
          QuelDateDeExo(VH^.EnCours.Code, ExoLog);
          // FQ 19497 - CA - Pas de log report à nouveau en clôture dans à nouveaux.
          // CPEnregistreLog('REPORT A NOUVEAU '+DateToStr(ExoLog.Deb));
          // GCO - 18/09/2006 - Traçage des évènements
          // L'exercice cloturé est devenu l'exercice précédent.    
          { BVE 29.08.07 : Mise en place d'un nouveau tracage }
{$IFNDEF CERTIFNF}
          CPEnregistreLog('CLOEXO ' + VH^.Precedent.Code);
{$ELSE}
          CPEnregistreJalEvent('CEX','Clôture exercice','CLOEXO ' + VH^.Precedent.Code);
{$ENDIF}
          //06/12/2006 YMO Norme NF 203 : interdiction d'annuler en cascade
          vAnnulClo:=GetParamSocSecur('SO_CPANNULCLO', False);
          SetParamSoc('SO_CPANNULCLO',Max(vAnnulClo-1,0));

        end;
      end;
    end else bErreur := True;
  end
  else
    bErreur := True;
  SetControlEnabled('BAUDIT', bErreur);
  SetControlEnabled('BSANS', bErreur);
  SetControlEnabled('BAVEC', bErreur);
end;

function TOF_CLOTUREDEF.OuvertureExerciceSuivant: boolean;
var
  ExoSuivant: TExoDate;
  TExercice : TOB;
begin
  // Si pas de création d'exercice suivant
  if m_bOuvrirSuivant then
  begin
    Result :=
      (ExecuteSQL('UPDATE EXERCICE SET EX_ETATCPTA="OUV" WHERE EX_DATEDEBUT="' +
      USDateTime(VH^.Encours.Fin + 1) + '" AND EX_ETATCPTA="NON"') = 1);
    ChargeMagExo(False);
    exit;
  end;
  if not m_bCreationSuivant then
  begin
    Result := (VH^.Suivant.Code <> '');
    exit;
  end;
  // Si création d'exercice suivant
  ExoSuivant.Deb := StrToDate(GetControlText('SUIVANTDEB'));
  ExoSuivant.Fin := StrToDate(GetControlText('SUIVANTFIN'));
  { On vérifie l'absence d'un tel exercice }// CA - 22/04/2003 ( FQ 10410 )
  if not ExisteSQL('SELECT * FROM EXERCICE WHERE EX_DATEDEBUT="' +
    USDateTime(ExoSuivant.Deb) + '"') then
  begin
    TExercice := TOB.Create ('EXERCICE',nil,-1);
    try
      TExercice.PutValue('EX_EXERCICE',TrouveNewCodeExo);
      TExercice.PutValue('EX_LIBELLE', LibelleExerciceDefaut(ExoSuivant.Deb, ExoSuivant.Fin, False));
      TExercice.PutValue('EX_ABREGE',LibelleExerciceDefaut(ExoSuivant.Deb, ExoSuivant.Fin, True));
      TExercice.PutValue('EX_DATEDEBUT',ExoSuivant.Deb);
      TExercice.PutValue('EX_DATEFIN',ExoSuivant.Fin);
      TExercice.PutValue('EX_ETATCPTA','OUV');
      TExercice.PutValue('EX_ETATBUDGET','OUV');
      TExercice.PutValue('EX_ETATADV','NON');
      TExercice.PutValue('EX_ETATAPPRO','NON');
      TExercice.PutValue('EX_ETATPROD','NON');
      TExercice.PutValue('EX_VALIDEE','------------------------');
      TExercice.PutValue('EX_DATECUM', 2);
      TExercice.PutValue('EX_DATECUMRUB', 2);
      TExercice.PutValue('EX_DATECUMBUD', 2);
      TExercice.PutValue('EX_DATECUMBUDGET',2);
      TExercice.PutValue('EX_NATEXO','');
      TExercice.InsertDB(nil);
    finally
      TExercice.Free;
    end;
  end;
  {$IFDEF EAGLCLIENT}
  AvertirCacheServer('EXERCICE');
  {$ENDIF}
  ChargeMagExo(False);
// CA - 26/04/2005 - Inutile car fait en fin de traitement et pose des problèmes parfois en clôture sans à-nouveau.
//  AvertirMultiTable('TTEXERCICE');
  if VH^.Suivant.Code <> '' then
    SetControlEnabled('SUIVANTFIN', False);
  Result := (VH^.Suivant.Code <> '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 04/04/2002
Modifié le ... :   /  /
Description .. : Récupère la date de début et de fin de l'exercice suivant.
Suite ........ : Si l'exercice suivant est ouvert, on récupère les infos dans
Suite ........ : VH^.Suivant.
Suite ........ : Si EX_ETATCPTA est à NON, on va chercher les infos
Suite ........ : dans la base.
Mots clefs ... :
*****************************************************************}

function TOF_CLOTUREDEF.GetInfoExerciceSuivant(var DateDeb: TDateTime;
  var DateFin: TDateTime): boolean;
var
  Q: TQuery;
begin
  Result := False;
  if VH^.Suivant.Code <> '' then
  begin
    DateDeb := VH^.Suivant.Deb;
    DateFin := VH^.Suivant.Fin;
    Result := True;
  end
  else
  begin
    Q := OpenSQL('SELECT * FROM EXERCICE WHERE EX_DATEDEBUT="' +
      USDateTime(VH^.Encours.Fin + 1) + '"', True);
    if not Q.Eof then
    begin
      if Q.FindField('EX_ETATCPTA').AsString = 'NON' then
      begin
        DateDeb := Q.FindField('EX_DATEDEBUT').AsDateTime;
        DateFin := Q.FindField('EX_DATEFIN').AsDateTime;
        m_bOuvrirSuivant := True;
        Result := True;
      end;
    end;
    Ferme(Q);
  end;
end;

function TOF_CLOTUREDEF.AucuneEcriture(Exo: string): boolean;
begin
  Result := not ExisteSQL('SELECT E_EXERCICE FROM ECRITURE WHERE E_EXERCICE="' +
    Exo + '"');
  if Result then
    PGIBox('Aucune écriture sur cet exercice. Clôture impossible.',
      ECRAN.Caption);
end;

initialization
  registerclasses([TOF_CLOTUREDEF]);
end.

