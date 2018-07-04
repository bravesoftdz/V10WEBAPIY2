{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 16/04/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMCLOTURE ()
Mots clefs ... : TOF;AMCLOTURE
Suite          : BTY 10/05 FQ 17085 Annulation de clôture, opérations de modif bases MBA/MB2 (et non MPA)
Suite......... : MBO 10/11/2005 : réécriture de la dépréciation d'actif sur IMMO à partir de
                 IMMOLOG - DPR si dépréciation saisie au cours de l'exercice clôturé
Suite          : BTY 11/05 Annulation de clôture, opérations de changement de plan CD2/CD3 en plus de CDM
Suite          : BTY 11/05 FQ 17086 Annulation de clôture, si plusieurs opérations de même type dans l'exercice,
Suite          : on a un message Annulation de clôture impossible ou une Erreur SQL générale
Suite          : PGR - 12/12/2005 - Annulation de clôture, ajout pour changement de méthode CDA
Suite..........: FQ 17215 - TGA 21/12/2005 - GetParamSoc => GetParamSocSecur
Suite......... : BTY - 01/06 FQ 17259 Nouveau top dépréciation dans IMMO
Suite......... : MVG - 31/03/06 MVG GETPARAMSOCSECUR de SO_DATECLOTUREIMMO à initialiser à idate1900 demande Y. PELUD
Suite......... : BTY - 04/06 FQ 17516 Nouveau top changement de regroupement dans IMMO
Suite......... : MBO - 04/06 FQ 17923 prise en compte nouvelle opération modif date de mise en service
                                      lors de l'annulation de la cloture
Suite......... : BTY - 05/06 FQ 18211 Positionner les indicateurs de modif de compta de PARAMSOC
Suite......... : MBO - 09/06 - Prise en compte de l'opération saisie d'une prime d'équipement PRI pour i_operation
Suite......... : MBO - 10/06 - Prise en compte de l'opération saisie d'une prime d'équipement PRI pour i_operation
Suite......... : BTY - 10/06 - Prise en compte de l'opération réduction de prime d'équipement RPR
Suite......... : BTY - 11/06 - Prise en compte de la réduction de subvention d'investissement RSB
Suite......... : BTY - 11/06 FQ 19193 CB ou LOC ayant subi une levée d'option ou suspendu restait dans l'historique
Suite......... : MBO - 18/04/2007 - restituer le code immo remplacée si opération remplacement composant
Suite......... : MBO - 22/06/2007 - correction plantage en cwas en annulation
*****************************************************************}
Unit AMCLOTURE_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uTOB,
     ParamSoc,
     ImEnt
//     uLibExercice,
//     utilpgi
     ;

Type
  TOF_AMCLOTURE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure AnnulationCloture;
    function JaiLeDroitAnnulationCloture : integer;
  end ;

procedure AMLanceFiche_AnnulationCloture;

Implementation

const MSG_CLOTURE : array [1..4] of string = (
        {1} 'Aucune clôture n''a été effectuée sur ce dossier',
        {2} 'Vous devez effectuer auparavant la suppression du journal d''à-nouveau.',
        {3} 'Des opérations sont présentes sur l''exercice.',
        {4} 'Des immobilisations ont été créées sur l''exercice en cours alors que leur date d''achat correspond à un exercice antérieur.#10#13 Veuillez les supprimer pour accéder à l''annulation de clôture.'
        );

procedure AMLanceFiche_AnnulationCloture;
begin
  if GetParamSocSecur('SO_EXOCLOIMMO','')<>'' then
  begin
    if not ImBlocageMonoPoste(True) then exit;
    try
      V_PGI.ZoomOLE := True;
      AGLLanceFiche('AM', 'AMCLOTURE', '', '','');
      V_PGI.ZoomOLE := False;
    finally
      ImDeblocageMonoPoste(True);
    end;
  end
  else PGIBox ('Aucune information de clôture n''a été trouvée sur ce dossier.','Annulation de clôture');
end;

procedure TOF_AMCLOTURE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMCLOTURE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AMCLOTURE.OnUpdate ;
var iJaiLeDroit : integer;
begin
  Inherited ;
  iJaiLeDroit := JaiLeDroitAnnulationCloture;
  if ( iJaiLeDroit = 0 ) then
  begin
    if (PGIAsk('Confirmez-vous l''annulation de la clôture des immobilisations ?')=mrYes) then
    begin
      if (Transactions(AnnulationCloture,1) <> oeOK) then
        PGIBox ('Annulation de clôture impossible.')
      else
      begin
      PGIInfo('Annulation de clôture réussie.',Ecran.Caption);
      // FQ 18211 Positionner les indicateurs de modif
      VHImmo^.ChargeOBImmo := True;
      ImMarquerPublifi(True);
      end;
    end;
  end else PGIBox('L''annulation de clôture est impossible. #10#13'+MSG_CLOTURE[iJaiLeDroit], ECRAN.Caption);
end ;

procedure TOF_AMCLOTURE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AMCLOTURE.OnArgument (S : String ) ;
var DateDebutExo, DateFinExo : TDateTime;
  i : integer;
begin
  Inherited ;
  SetControlVisible('TRAITEMENTENCOURS',False);
  if GetParamSocSecur('SO_EXOCLOIMMO','')<>'' then
  begin
    SetControlText('DATECLOTURE',DateToStr(GetParamSocSecur('SO_DATECLOTUREIMMO',idate1900)));
    DateDebutExo := iDate1900;
    DateFinExo := iDate1900;
    for i:=1 to 20 do
    begin
      if VHImmo^.Exercices[i].Code = GetParamSocSecur('SO_EXOCLOIMMO','') then
      begin
        DateDebutExo := VHImmo^.Exercices[i].Deb;
        DateFinExo := VHImmo^.Exercices[i].Fin;
        break;
      end;
    end;
    SetControlText('EXERCICECLOTURE',DateToStr(DateDebutExo)+' au '+DateToStr(DateFinExo));
  end;
end ;

procedure TOF_AMCLOTURE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AMCLOTURE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMCLOTURE.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 20/01/2005
Modifié le ... :   /  /    
Description .. : FQ 15303 - CA - 20/01/2005 - Annulation de clôture 
Suite ........ : interdite si des immobilisations reprises sont présentes sur 
Suite ........ : l'exercice.
Mots clefs ... : 
*****************************************************************}
function TOF_AMCLOTURE.JaiLeDroitAnnulationCloture : integer;
var ExoCloImmo : string;
begin
  Result := 0;
  ExoCloImmo := GetParamSocSecur('SO_EXOCLOIMMO','');
  { Il doit y avoir quelque chose dans SO_EXOCLOIMMO }
  if  ExoCloImmo = '' then Result := 1
  { L'exercice clos en immobilisations doit être ouvert en comptabilité }
  else if not ExisteSQL('SELECT EX_ETATCPTA FROM EXERCICE WHERE EX_EXERCICE="'+ExoCloImmo+'" AND EX_ETATCPTA="OUV"') then Result := 2
  { Il ne doit pas y avoir d''opérations en cours }
  else if ExisteSQL('SELECT IL_IMMO FROM IMMOLOG WHERE IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Fin+1)+'" AND IL_TYPEOP<>"ACQ"') then Result := 3
  { il ne doit pas y avoir d'immobilisations saisies sur l'exercice avec un reprise - FQ 15303}
  else if ExisteSQL('SELECT I_IMMO FROM IMMO WHERE (I_REPRISEECO<>0 OR I_REPRISEFISCAL<>0) AND I_ETAT="OUV"') then Result := 4;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 11/01/2005
Modifié le ... :   /  /
Description .. : FQ 15238 - CA - 11/01/2005 - Ne pas enlever de
Suite ........ : l'historique des immobilisations cédées antérieurement.
Mots clefs ... :
*****************************************************************}
procedure TOF_AMCLOTURE.AnnulationCloture;
var Q : TQuery;
    stCode, stEtat : string;
    stUpdate : string;
    i, iPlanActif : integer;
    TLog : TOB;
    bSortie, bCloture, bMemCloture, bLevee : boolean;
    bAcquisitionAnnee : boolean;
    DateDebutExo, DateFinExo,DateSuspCB : TDateTime;
    iNombreImmoTotal, iCompteurImmo : integer;
    cherche : string;
    // ajout mbo pour prime
    bPrime : boolean;
    bSbv, bReductionPrime, bReductionSBV, CBouLOC : boolean;
    remplace : string;  // ajout mbo pour remplacement de composant
    Q2 : TQuery;
    requetembo : string;
begin
  bMemCloture := False;
  bPrime := false;
  bSbv := false;
  bReductionPrime := False;
  bReductionSBV := False;
  //CBouLOC := False; //XVI Conseil Compile...
  DateDebutExo := VHImmo^.EnCours.Deb;
  DateFinExo := VHImmo^.EnCours.Fin;
  SetControlVisible('TRAITEMENTENCOURS',True);
  iCompteurImmo := 0;
  { Calcul du nombre de fiches à traiter }
  Q := OpenSQL('SELECT COUNT(I_IMMO) NOMBREIMMO FROM IMMO WHERE I_DATEPIECEA<="'+USDateTime(DateFinExo)+'"',True);
  iNombreImmoTotal := Q.FindField('NOMBREIMMO').AsInteger;
  Ferme (Q);

  { 1 - Suppression des enregistrements de IMMOLOG }
  SetControlText ('TRAITEMENTENCOURS', 'Suppression de l''historique en cours ...');

  //ExecuteSQL ('DELETE FROM IMMOLOG WHERE (IL_DATEOP>"'+USDateTime(DateFinExo)+'" AND IL_TYPEOP<>"ACQ") OR '+
  //      '(IL_DATEOP="'+USDateTime(DateFinExo)+'" AND IL_TYPEOP="CLO")');

  { Pour chaque immobilisation concernée par la clôture }
  TLog := TOB.Create ('',nil,-1);
  try
    // FQ 19193
    //Q := OpenSQL ('SELECT I_IMMO,I_LIBELLE, I_ETAT FROM IMMO WHERE I_DATEPIECEA<="'+USDateTime(DateFinExo)+'"', True );
    Q := OpenSQL ('SELECT I_IMMO,I_LIBELLE,I_ETAT,I_NATUREIMMO,I_DATESUSPCB FROM IMMO WHERE I_DATEPIECEA<="'+USDateTime(DateFinExo)+'"', True );
    try
      while not Q.Eof do
      begin
        Inc(iCompteurImmo);
        SetControlText ('TRAITEMENTENCOURS', Q.FindField('I_LIBELLE').AsString+' ...'+IntToStr(Trunc((iCompteurImmo/iNombreImmoTotal)*100))+' %...');
        Application.ProcessMessages;
        stCode := Q.FindField('I_IMMO').AsString;
        stEtat := Q.FindField('I_ETAT').AsString;
        // FQ 19193
        DateSuspCB := Q.FindField('I_DATESUSPCB').AsDateTime;
        // FQ 19193
        CBouLOC := ((Q.FindField('I_NATUREIMMO').AsString = 'CB') or (Q.FindField('I_NATUREIMMO').AsString = 'LOC'));
        iPlanActif := 0;
        stUpdate := '';
        bCloture := False;
        bSortie := False;
        bLevee := False;  // FQ 19193
        bAcquisitionAnnee := False;

        // on recupere le code mutation de immolog de cloture avant de détruire l'enreg de cloture
        // au cas ou cette immo est issue d'un remplacement de composant
        Q2 := OpenSQL ('SELECT IL_CODEMUTATION FROM IMMOLOG WHERE IL_DATEOP="'+USDateTime(DateFinExo)+'" AND IL_TYPEOP="CLO" ' +
                   'AND IL_IMMO = "'+ stCode +'"', true);
        // ajout mbo 22.06.07 suite plantage en cwas
        remplace := '';
        if not Q2.Eof then
           remplace := Q2.FindField('IL_CODEMUTATION').AsString;

        Ferme(Q2);

        requetembo := 'DELETE FROM IMMOLOG WHERE ((IL_DATEOP>"'+USDateTime(DateFinExo)+'" AND IL_TYPEOP<>"ACQ") OR '+
                   '(IL_DATEOP="'+USDateTime(DateFinExo)+'" AND IL_TYPEOP="CLO")) ' +
                   'AND (IL_IMMO = "'+ stCode +'")' ;

        ExecuteSQL ('DELETE FROM IMMOLOG WHERE ((IL_DATEOP>"'+USDateTime(DateFinExo)+'" AND IL_TYPEOP<>"ACQ") OR ' +
                   '(IL_DATEOP="'+USDateTime(DateFinExo)+'" AND IL_TYPEOP="CLO")) ' +
                   'AND (IL_IMMO = "'+ stCode +'")');

        { 2 - Mise à jour de la fiche}
        TLog.LoadDetailFromSQL('SELECT IL_IMMO,IL_TYPEOP,IL_DATEOP,IL_PLANACTIFAP, IL_MONTANTDOT FROM IMMOLOG WHERE IL_IMMO="'+stCode+'" ORDER BY IL_DATEOP',False);
        if stEtat<>'FER' then stEtat := 'OUV';
        for i:=0 to TLog.Detail.Count - 1 do
        begin
          { Le plan actif est-il le plus grand ? }
          if (TLog.Detail[i].GetValue('IL_PLANACTIFAP') > iPlanActif) then
            iPlanActif := TLog.Detail[i].GetValue('IL_PLANACTIFAP');

          { Est-ce qu'une clôture a déjà eu lieu ? }
          if (TLog.Detail[i].GetValue('IL_TYPEOP')='CLO') then  bCloture := True;

          { Est-ce que l'immobilisation a été sortie durant l'exercice }
          if ((TLog.Detail[i].GetValue('IL_DATEOP')>=DateDebutExo) and (TLog.Detail[i].GetValue('IL_DATEOP')<=DateFinExo))
                  and (TLog.Detail[i].GetValue('IL_TYPEOP')='CES') then bSortie := True;

          // FQ 19193 Identifier immo ayant subi une levée d'option dans l'exercice
          if ((TLog.Detail[i].GetValue('IL_DATEOP')>=DateDebutExo) and (TLog.Detail[i].GetValue('IL_DATEOP')<=DateFinExo))
                  and (TLog.Detail[i].GetValue('IL_TYPEOP')='LEV') then bLevee := True;

          // ajout mbo pour prime
          if ((TLog.Detail[i].GetValue('IL_DATEOP')>=DateDebutExo) and (TLog.Detail[i].GetValue('IL_DATEOP')<=DateFinExo))
                  and (TLog.Detail[i].GetValue('IL_TYPEOP')='PRI') then bPrime := True
          else
             bPrime := false;

          // ajout mbo pour subvention
          if ((TLog.Detail[i].GetValue('IL_DATEOP')>=DateDebutExo) and (TLog.Detail[i].GetValue('IL_DATEOP')<=DateFinExo))
                  and (TLog.Detail[i].GetValue('IL_TYPEOP')='SBV') then bSBV := True
          else
             bSBV := false;

          // BTY 10/06 Opération de réduction de prime
          bReductionPrime := ((TLog.Detail[i].GetValue('IL_DATEOP')>=DateDebutExo) and
                              (TLog.Detail[i].GetValue('IL_DATEOP')<=DateFinExo) and
                              (TLog.Detail[i].GetValue('IL_TYPEOP')='RPR'));
          // BTY 11/06 Opération de réduction de subvention
          bReductionSBV := ((TLog.Detail[i].GetValue('IL_DATEOP')>=DateDebutExo) and
                             (TLog.Detail[i].GetValue('IL_DATEOP')<=DateFinExo) and
                             (TLog.Detail[i].GetValue('IL_TYPEOP')='RSB'));

          { Quels ont été les opérations effectuées durant l'exercice }
          // BTY 11/05 FQ 17086 Attention de ne pas répéter le même terme dans stUpdate sinon requête fausse et Erreur SQL générale
          if ((TLog.Detail[i].GetValue('IL_DATEOP')>=DateDebutExo) and (TLog.Detail[i].GetValue('IL_DATEOP')<=DateFinExo)) then
          begin
            if (TLog.Detail[i].GetValue('IL_TYPEOP')='MUT') then
            begin
                //stUpdate := stUpdate+'I_OPEMUTATION="X",'
                cherche := 'I_OPEMUTATION="X"';
                if POS(cherche,stUpdate) = 0 then stUpdate := stUpdate+'I_OPEMUTATION="X",';
            end
            else if (TLog.Detail[i].GetValue('IL_TYPEOP')='ECL') then
            begin
                //stUpdate := stUpdate+'I_OPEECLATEMENT="X",'
                cherche := 'I_OPEECLATEMENT="X"';
                if POS(cherche,stUpdate) = 0 then stUpdate := stUpdate+'I_OPEECLATEMENT="X",';
            end
            //else if (TLog.Detail[i].GetValue('IL_TYPEOP')='CDM') then stUpdate := stUpdate+'I_OPECHANGEPLAN="X",'
            // BTY 11/05 Opérations de révision de plan avec plan futur sur VNC et ajout plan fiscal CRC 2002-10
            // mbo 04/2006 fq 17923
            else if (TLog.Detail[i].GetValue('IL_TYPEOP')='CDM') or
                    (TLog.Detail[i].GetValue('IL_TYPEOP')='CD2') or
                    (TLog.Detail[i].GetValue('IL_TYPEOP')='CD3') then
            begin
                //stUpdate := stUpdate+'I_OPECHANGEPLAN="X",'
                cherche := 'I_OPECHANGEPLAN="X"';
                if POS(cherche,stUpdate) = 0 then stUpdate := stUpdate+'I_OPECHANGEPLAN="X",';
            end
            // mbo 04/2006 fq 17923
            else if (TLog.Detail[i].GetValue('IL_TYPEOP')='MMS') then
            begin
                //stUpdate := stUpdate+'I_OPECHANGEPLAN="X",'
                cherche := 'I_OPECHANGEPLAN="X"';
                if POS(cherche,stUpdate) = 0 then stUpdate := stUpdate+'I_OPECHANGEPLAN="X",';
            end
            else if (TLog.Detail[i].GetValue('IL_TYPEOP')='ELC') then
            begin
                //stUpdate := stUpdate+'I_OPECHANGEPLAN="X",'
                cherche := 'I_OPECHANGEPLAN="X"';
                if POS(cherche,stUpdate) = 0 then stUpdate := stUpdate+'I_OPECHANGEPLAN="X",';
            end
            else if (TLog.Detail[i].GetValue('IL_TYPEOP')='ELE') then
            begin
                //stUpdate := stUpdate+'I_OPECHANGEPLAN="X",'
                cherche := 'I_OPECHANGEPLAN="X"';
                if POS(cherche,stUpdate) = 0 then stUpdate := stUpdate+'I_OPECHANGEPLAN="X",';
            end
            else if (TLog.Detail[i].GetValue('IL_TYPEOP')='LIE') then
            begin
                //stUpdate := stUpdate+'I_OPELIEUGEO="X",'
                cherche := 'I_OPELIEUGEO="X"';
                if POS(cherche,stUpdate) = 0 then stUpdate := stUpdate+'I_OPELIEUGEO="X",';
            end
            else if (TLog.Detail[i].GetValue('IL_TYPEOP')='ETA') then
            begin
                //stUpdate := stUpdate+'I_OPEETABLISSEMENT="X",'
                cherche := 'I_OPEETABLISSEMENT="X"';
                if POS(cherche,stUpdate) = 0 then stUpdate := stUpdate+'I_OPEETABLISSEMENT="X",';
            end
            else if (TLog.Detail[i].GetValue('IL_TYPEOP')='LEV') then
            begin
                //stUpdate := stUpdate+'I_OPELEVEEOPTION="X",'
                cherche := 'I_OPELEVEEOPTION="X"';
                if POS(cherche,stUpdate) = 0 then stUpdate := stUpdate+'I_OPELEVEEOPTION="X",';
            end
            // BTY 04/06 Changement de regroupement
            else if (TLog.Detail[i].GetValue('IL_TYPEOP')='REG') then
            begin
                cherche := 'I_OPEREG="X"';
                if POS(cherche,stUpdate) = 0 then stUpdate := stUpdate+'I_OPEREG="X",';
            end
            // BTY 10/05 MBA et non MPA
            // BTY 10/05 Avec un plan fiscal CRC 2002-10 la modif bases est MB2
            // FQ 17085 else if (TLog.Detail[i].GetValue('IL_TYPEOP')='MPA') then stUpdate := stUpdate+'I_OPEMODIFBASES="X",'
            else if (TLog.Detail[i].GetValue('IL_TYPEOP')='MBA') or
                    (TLog.Detail[i].GetValue('IL_TYPEOP')='MB2') then
            begin
                    //stUpdate := stUpdate+'I_OPEMODIFBASES="X",'
                    cherche := 'I_OPEMODIFBASES="X"';
                    if POS(cherche,stUpdate) = 0 then stUpdate := stUpdate+'I_OPEMODIFBASES="X",';
            end
            // ajout mbo pour depreciation d'actif 10.11.05
            else if (TLog.Detail[i].GetValue('IL_TYPEOP')='DPR') then
                 begin
                 stUpdate := stUpdate+'I_REVISIONECO='+ StrFPoint(TLog.Detail[i].GetValue('IL_MONTANTDOT'))+',';
                 cherche := 'I_OPECHANGEPLAN="X"';
                 if POS(cherche,stUpdate) = 0 then
                    stUpdate := stUpdate+'I_OPECHANGEPLAN="X",';
                 // BTY 01/06 FQ 17259 Nouveau top de dépréciation
                 cherche := 'I_OPEDEPREC="X"';
                 if POS(cherche,stUpdate) = 0 then
                    stUpdate := stUpdate+'I_OPEDEPREC="X",';
                 end
             //PGR - 12/12/2005 - Ajout pour changement de méthode
             else if (TLog.Detail[i].GetValue('IL_TYPEOP')='CDA') then
                 begin
                 cherche := 'I_OPECHANGEPLAN="X"';
                 if POS(cherche,stUpdate) = 0 then
                    stUpdate := stUpdate+'I_OPECHANGEPLAN="X",';
                 end
            else if (TLog.Detail[i].GetValue('IL_TYPEOP')='ACQ') then bAcquisitionAnnee := True;
          end;
        end;
        TLog.ClearDetail;


        if bSortie then stUpdate := stUpdate+'I_OPECESSION="X",';
        if stUpdate<>'' then stUpdate:=','+stUpdate;

        // MODIF MBO POUR PRIME if stUpdate <> '' then stUpdate := stUpdate + 'I_OPERATION="X" ';
        // BTY Réduction prime
        // BTY Réduction subvention
        if stUpdate <> ''  then
           stUpdate := stUpdate + 'I_OPERATION="X" '
        else if (bPrime or bSBV or bReductionPrime or bReductionSBV) then
                 stUpdate := ',' + 'I_OPERATION="X" ';

        // ajout mbo pour remplacement de composant
        if TRIM(remplace) <> '' then
        begin
           stUpdate := stUpdate+',' + 'I_REMPLACE="'+ remplace + '",';
           stUpdate := stUpdate+'I_TYPER="IRM"';
        end;
        //

        { Calcul du champ I_ETAT }
        if (stEtat='FER') and bSortie and bCloture then stEtat := 'CLO'
        else
        if (stEtat='FER') and bLevee and bCloture then stEtat := 'CLO'   // FQ 19193
        else if (stEtat='FER') and bSortie and (not bCLoture) then stEtat := 'OUV'
        else if (stEtat='FER') and bLevee and (not bCLoture) then stEtat := 'OUV' // FQ 19193
        else if (stEtat='FER') and CBouLOC and (DateSuspCB<>iDate1900) and (not bCloture) then stEtat := 'OUV' // FQ 19193
        else if (stEtat<>'FER') and bCloture then stEtat := 'CLO'
        { FQ 15238 - on ne repasse à "OUV" que les acquisitions de l'année }
        else if bAcquisitionAnnee then stEtat := 'OUV';
        ExecuteSQL ('UPDATE IMMO SET I_ETAT="'+stEtat+'", I_PLANACTIF='+IntToStr(iPlanActif)+stUpdate+' WHERE I_IMMO="'+stCode+'"');
        { 3 - Suppression des plans de clôture et d'opérations après clôture }
        ExecuteSQL ('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+stCode+'" AND IA_NUMEROSEQ>'+IntToStr(iPlanActif));
        if bCloture then bMemCloture := True;
        Q.Next;
      end;
    finally
      Ferme (Q);
    end;
  finally
    TLog.Free;
  end;
  { 4 - Mise à jour des paramsocs }
  if bMemCloture then
  begin
    SetParamSoc('SO_DATECLOTUREIMMO',DateFinExo);
    if VHImmo^.Precedent.Code <> '' then SetParamSoc('SO_EXOCLOIMMO',VHImmo^.Precedent.Code)
    else SetParamSoc('SO_EXOCLOIMMO','');
  end else
  begin
    SetParamSoc('SO_DATECLOTUREIMMO',iDate1900);
    SetParamSoc('SO_EXOCLOIMMO','');
  end;
  SetControlText ('TRAITEMENTENCOURS', 'Traitement terminé.');
end;

Initialization
  registerclasses ( [ TOF_AMCLOTURE ] ) ;
end.
