{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 06/08/2001
Modifié le ... :   /  /
Description .. : Unit de génération automatique des acomptes d'un mois sur
Suite ........ : l'autre. Ne reprend que les colonnes à reporter.
Suite ........ : Les acomptes sont dans la table HISTOSAISRUB
Mots clefs ... : PAIE;PGACOMPTE
*****************************************************************}
{
PT1 : 23/04/2002 V571 PH initialisation du champ PSD_RIBSALAIRE
PT2 : 03/06/2002 V582 PH Gestion historique des évènements
PT3 : 08/08/2002 V585 PH Correction requete si selection Etablissement
PT4 : 08/10/2004 V_50 PH FQ 11678 Ne générer un acompte en auto sur un salarié sorti
PT5 : 10/11/2004 V_60 PH FQ 11678 suite Prise en compte date de sortie non renseignée
PT6 : 13/04/2006 V_65 PH FQ 12501 Message en fin de traitement
PT7 : 10/01/2008 V_802 FLO FQ 14798 Ne pas forcément supprimer les acomptes existants si date différente
}
unit UTofPG_GENERACOMPTE;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DBGrids,
{$ENDIF}
  Grids, HCtrls, HEnt1, HMsgBox, UTOF, UTOB, Vierge, P5Util, P5Def, AGLInit, EntPaie,
  PgOutils2;

type
  TOF_PG_GENERACOMPTE = class(TOF)
  private
    VCbxLAnnee: THValComboBox;
    VCbxLMois: THValComboBox;
    VCbxLEtab: THValComboBox;
    Etab: string;
    procedure LanceGenerAcompte(TOB_Acompte: TOB; DDPR, DFPR: TDateTime);
    procedure GenerAcompteEnter(Sender: TObject);
    function AnnuleAcompte(Chaine: string): boolean;
  public
    procedure OnClose; override;
    procedure OnArgument(Arguments: string); override;
  end;

implementation


function TOF_PG_GENERACOMPTE.AnnuleAcompte(Chaine: string): boolean;
begin
  result := TRUE;
  try
    BeginTrans;
    ExecuteSQL(chaine);
    CommitTrans;
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de la suppression des acomptes', Ecran.caption);
    result := FALSE;
  end;

end;


procedure TOF_PG_GENERACOMPTE.GenerAcompteEnter(Sender: TObject);
var
  Mois, Annee: WORD;
  DD, DF, DDPR, DFPR: TDateTime;
  St, St1: string;
  TOB_Acompte: TOB;
  Q: TQuery;
  LaSuite: Boolean;
  i,Rep : Integer; //PT7
begin
  Rep := MrYes; //PT7
  if VCbxLMois <> nil then Mois := StrToInt(VCBXLMois.Value)
  else exit;
  if VCbxLAnnee <> nil then Annee := StrToInt(VCBXLAnnee.Text)
  else exit;
  if VCbxLEtab <> nil then Etab := VCbxLEtab.Value;
  DD := EncodeDate(Annee, Mois, 01);
  DF := FINDEMOIS(DD);
  if Etab = '' then
    St := ' SELECT PSD_SALARIE FROM HISTOSAISRUB WHERE PSD_ORIGINEMVT="ACP" AND PSD_DATEDEBUT >="' +
      UsDateTime(DD) + '" AND PSD_DATEFIN<="' + UsDateTime(DF) + '"'
  else
  begin
    st := 'SELECT PSD_SALARIE FROM HISTOSAISRUB WHERE PSD_ETABLISSEMENT = "' +
      Etab + '" AND PSD_ORIGINEMVT="ACP" AND PSD_DATEDEBUT >="' +
      UsDateTime(DD) + '" AND PSD_DATEFIN<="' + UsDateTime(DF) + '"';
  end;
  Q := OpenSql(St, TRUE);
  LaSuite := TRUE;
  // on va chercher à vérifier si on pas déjà généré ou saisi des acomptes sur la même période pour le même établisseement
  if not Q.EOF then
  begin
  	//PT7 - Début
    (*St := 'Attention, vous avez déjà des acomptes établis #13#10 entre le ' + DateToStr(DD) + 'et le ' + DateToStr(DF) +
      '.#13#10Voulez vous continuer ?';
    if PGIAsk(St, Ecran.Caption) = MrYes then
    begin
      if PGIAsk('Ces acomptes vont être détruits pour être régénérés.#13#10Voulez-vous continuer ?', Ecran.Caption) = MrYes then
      begin
    *)
    St := 'Attention, vous avez déjà des acomptes établis entre le ' + DateToStr(DD) + ' et le ' + DateToStr(DF) +'.'+
          '#13#10Ils doivent être détruits pour être régénérés.'+
          '#13#10#13#10Voulez vous supprimer tous les acomptes de la période?';
    Rep := PGIAskCancel(St, Ecran.Caption);
    If Rep = MrYes Then  
    Begin
    //PT7 - Fin
        if etab = '' then
          St1 := 'DELETE FROM HISTOSAISRUB WHERE PSD_ORIGINEMVT="ACP" AND PSD_DATEDEBUT >="' +
            UsDateTime(DD) + '" AND PSD_DATEFIN<="' + UsDateTime(DF) + '"'
        else st1 := 'DELETE FROM HISTOSAISRUB WHERE PSD_ORIGINEMVT="ACP" AND PSD_DATEDEBUT >="' +
          UsDateTime(DD) + '" AND PSD_DATEFIN<="' + UsDateTime(DF) + '"' +
            ' AND PSD_ETABLISSEMENT="' + Etab + '"';
        LaSuite := AnnuleAcompte(St1);
    End 
    Else If Rep = MrNo Then
    Begin
    	PGIInfo(TraduireMemoire('Seuls les acomptes déjà reportés du mois précédent seront supprimés.'),TraduireMemoire('Information'));
    (* //PT7
      end
      else LaSuite := FALSE;*)
    end
    else LaSuite := FALSE;
  end;

  Ferme(Q);
  // Si OK on construit la liste des acomptes du mois precedent, en verifiant que le salarie est bien prevu sur l'etablissement demande
  // et non pas sur l'etablissement de l'acompte du mois precedant car il peut avoir changer d'etablissement dans le mois
  if LaSuite then
  begin
    TOB_Acompte := TOB.Create('Les Acomptes du mois', nil, -1);
    St := 'Voulez vous générer les acomptes du ' + DateToStr(DD) + ' au ' + DateToStr(DF) + '?';
    if PGIAsk(St, Ecran.Caption) = mrYes then
    begin
      DDPR := PLUSMOIS(DD, -1);
      DFPR := PLUSMOIS(DF, -1);
      // PT3 : 08/08/2002 V585 PH Correction requete si selection Etablissement
      if etab = '' then
        St := 'SELECT * FROM HISTOSAISRUB WHERE PSD_ORIGINEMVT="ACP" AND PSD_AREPORTER="OUI" AND PSD_DATEDEBUT >="' +
          UsDateTime(DDPR) + '" AND PSD_DATEFIN<="' + UsDateTime(DFPR) + '" ORDER BY PSD_ORIGINEMVT,PSD_SALARIE'
      else st := 'SELECT * FROM HISTOSAISRUB left join SALARIES ON PSA_SALARIE=PSD_SALARIE' +
        ' WHERE PSD_ORIGINEMVT="ACP" AND PSD_AREPORTER="OUI" AND PSD_DATEDEBUT >="' +
          UsDateTime(DDPR) + '" AND PSD_DATEFIN<="' + UsDateTime(DFPR) + '" AND PSA_ETABLISSEMENT = "' +
          Etab + '" ORDER BY PSD_ORIGINEMVT,PSD_SALARIE';
      Q := OpenSql(St, TRUE);
      if not Q.EOF then
      begin
        TOB_Acompte.LoadDetailDB('HISTOSAISRUB', '', '', Q, False);
        
        //PT7 - Début
        // Suppression des acomptes déjà générés si l'utilisateur n'a pas voulu tout supprimer
        If Rep = MrNo Then
        Begin
        	St := '';
        	For i := 0 To TOB_Acompte.Detail.Count-1 Do
        	Begin
        		if i > 0 Then St := St + ' OR ';
	        	St := St + '(PSD_SALARIE="'+TOB_Acompte.Detail[i].GetValue('PSD_SALARIE')+'" AND ' +
	        		  'PSD_DATEDEBUT="'+UsDateTime(PlusMois(TOB_Acompte.Detail[i].GetValue('PSD_DATEDEBUT'),1))+'" AND '+
	        		  'PSD_DATEFIN="'+UsDateTime(PlusMois(TOB_Acompte.Detail[i].GetValue('PSD_DATEFIN'),1))+'")';
        	End;
        	ExecuteSQL('DELETE FROM HISTOSAISRUB WHERE '+St);
        End;
        //PT7 - Fin
        
        LanceGenerAcompte(TOB_Acompte, DDPR, DFPR);
      end
      else PGIBox('Vous n''avez pas d''acompte à reporter du mois précédent', Ecran.Caption);
      Ferme(Q);
    end;
    FreeAndNil(TOB_Acompte);
  end;

end;

procedure TOF_PG_GENERACOMPTE.LanceGenerAcompte(TOB_Acompte: TOB; DDPR, DFPR: TDateTime);
var
  st, LeSal, PaiAcp: string;
  TOBA, TOBA_NEW, TOBA_PR, TOB_Sal, TS, TE, TEtab: TOB; // Tob
  DDA, DFA: TDateTime;
  //  retour: Boolean;
  Q: TQuery;
  St1: string;
  Trace: TStringList;
  Err: Boolean;
begin
  // PT2 : 03/06/2002 V582 PH Gestion historique des évènements
  Trace := TStringList.Create;
  Err := FALSE;

  TEtab := Tob.create('Les Etabs', nil, -1);
  st := 'SELECT * FROM ETABCOMPL';
  Q := OpenSql(st, TRUE);
  TEtab.LoadDetailDB('ETABCOMPL', '', '', Q, FALSE);
  Ferme(Q);
  TOBA := TOB.Create('Les acomptes du mois', nil, -1);
  TOBA_PR := TOB_Acompte.FindFirst([''], [''], FALSE);
  SourisSablier;
  // Il faut prendre les caractéristiques du salarie à savoir les champs CONFIDENTIEL, ETABLISSEMENT
  // au monent ou on genere les acomptes en automatique
  TOB_Sal := TOB.Create('Les Salaries', nil, -1);
  // DEB PT4
  if Etab = '' then st := 'SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PSA_PAIACOMPTE,PSA_TYPPAIACOMPT,PSA_CONFIDENTIEL,PSA_DATESORTIE FROM SALARIES ORDER BY PSA_SALARIE'
  else st := 'SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PSA_PAIACOMPTE,PSA_TYPPAIACOMPT,PSA_CONFIDENTIEL,PSA_DATESORTIE FROM SALARIES WHERE PSA_ETABLISSEMENT="' + Etab + '" ORDER BY PSA_SALARIE';
  Q := OpenSql(St, TRUE);
  TOB_Sal.LoadDetailDB('SALARIES', '', '', Q, False);
  Ferme(Q);
  // Fin Chargement des salaries concernes
  while TOBA_PR <> nil do
  begin
    LeSal := TOBA_PR.GetValue('PSD_SALARIE');
    TS := TOB_Sal.FindFirst(['PSA_SALARIE'], [LeSal], FALSE);
    if TS <> nil then // Salarie ok on va charger etablissement et confidentiel
    begin
      DDA := TOBA_PR.GetValue('PSD_DATEDEBUT');
      DFA := TOBA_PR.GetValue('PSD_DATEFIN');
      DDA := PLUSMOIS(DDA, 1);
      DFA := PLUSMOIS(DFA, 1);
      // PT5 Test date de sortie par rapport à iDate1900
      if (DDA < TS.GetValue('PSA_DATESORTIE')) or (TS.GetValue('PSA_DATESORTIE') <= IDate1900) then
      begin // FIN PT4
        TOBA_NEW := TOB.Create('HISTOSAISRUB', TOBA, -1);
        TOBA_NEW.putValue('PSD_ORIGINEMVT', 'ACP');
        TOBA_NEW.putValue('PSD_SALARIE', LeSal);
        TOBA_NEW.PutValue('PSD_DATEDEBUT', DDA);
        TOBA_NEW.PutValue('PSD_DATEFIN', DFA);
        TOBA_NEW.putValue('PSD_RUBRIQUE', VH_PAIE.PGRubAcompte);
        TOBA_NEW.putValue('PSD_ORDRE', 0);
        TOBA_NEW.putValue('PSD_LIBELLE', 'Acompte du ' + DateToStr(DDA));
        TOBA_NEW.PutValue('PSD_ETABLISSEMENT', TS.GetValue('PSA_ETABLISSEMENT'));
        TOBA_NEW.PutValue('PSD_CONFIDENTIEL', TS.GetValue('PSA_CONFIDENTIEL'));
        // PT1 : 23/04/2002 V571 PH initialisation du champ PSD_RIBSALAIRE au lieu de PSD_RIBSALARIE !!
        TOBA_NEW.PutValue('PSD_RIBSALAIRE', '');
        TOBA_NEW.PutValue('PSD_BANQUEEMIS', '');
        TOBA_NEW.PutValue('PSD_TOPREGLE', '-');
        TOBA_NEW.PutValue('PSD_DATEINTEGRAT', IDate1900);
        TOBA_NEW.PutValue('PSD_DATECOMPT', IDate1900);
        if TS.getvalue('PSA_TYPPAIACOMPT') = 'PER' then PaiAcp := TS.getvalue('PSA_PAIACOMPTE')
        else
        begin
          TE := TEtab.FindFirst(['ETB_ETABLISSEMENT'], [TS.GetValue('PSA_ETABLISSEMENT')], FALSE);
          if TE <> nil then PaiAcp := TE.GetValue('ETB_PAIACOMPTE');
        end;
        if PaiAcp <> '008' then TOBA_NEW.PutValue('PSD_DATEPAIEMENT', DDA)
        else TOBA_NEW.PutValue('PSD_DATEPAIEMENT', iDate1900);  //PT7
        TOBA_NEW.PutValue('PSD_TYPALIMPAIE', 'M');
        TOBA_NEW.PutValue('PSD_BASE', 0);
        TOBA_NEW.PutValue('PSD_TAUX', 0);
        TOBA_NEW.PutValue('PSD_COEFF', 0);
        TOBA_NEW.PutValue('PSD_MONTANT', TOBA_PR.GetValue('PSD_MONTANT'));
        TOBA_NEW.PutValue('PSD_AREPORTER', TOBA_PR.GetValue('PSD_AREPORTER'));
      end;
    end;
    TOBA_PR := TOB_Acompte.FindNext([''], [''], FALSE);
  end;
  SourisNormale;
  try
    BeginTrans;
    TOBA.InsertDB(nil, FALSE);
    CommitTrans;
    PgiBox ('Génération automatique des acomptes terminée', Ecran.Caption);//PT6
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de la validation des acomptes', Ecran.Caption);
    Err := TRUE;
  end;
  // PT2 : 03/06/2002 V582 PH Gestion historique des évènements

  St1 := '';
  if Err then st1 := 'Erreur de ';
  st := 'Génération automatique des acomptes ';
  Trace.add(St1 + St);
  if not Err then Trace.add(IntToStr(TOBA.detail.count - 1) + ' acomptes générées');
  if Err then CreeJnalEvt('001', '008', 'ERR', nil, nil, Trace)
  else CreeJnalEvt('001', '008', 'OK', nil, nil, Trace);
  Trace.free;
  // FIN PT2
  FreeAndNil(TOBA);
  FreeAndNil(TOB_Sal);
  FreeAndNil(TEtab);
end;

procedure TOF_PG_GENERACOMPTE.OnArgument(Arguments: string);
var
  MoisE, AnneeE, ComboExer: string;
  BtnVal: TToolbarButton97;
  DebExer, FinExer: TDateTime;
begin
  inherited;
  BtnVal := TToolbarButton97(GetControl('BValider'));
  if BtnVal <> nil then BtnVal.OnClick := GenerAcompteEnter;
//  if not BlocageMonoPoste(TRUE) then exit;
  VCbxLAnnee := THValComboBox(GetControl('VCBXANNEE'));
  VCbxLMois := THValComboBox(GetControl('VCBXMOIS'));
  VCbxLEtab := THValComboBox(GetControl('VCBETAB'));

  if RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer) = TRUE then
  begin
    if vcbxLAnnee <> nil then
    begin
      vcbxLAnnee.value := ComboExer;
    end;
  end;
  if VCbxLMois <> nil then VCbxLMois.Value := MoisE;
end;

procedure TOF_PG_GENERACOMPTE.OnClose;
begin
  DeblocageMonoPoste(TRUE);
end;

initialization
  registerclasses([TOF_PG_GENERACOMPTE]);
end.

