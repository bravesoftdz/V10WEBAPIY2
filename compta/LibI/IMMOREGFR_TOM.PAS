{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 04/10/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : IMMOREGFR (IMMOREGFR)
Mots clefs ... : TOM;IMMOREGFR
*****************************************************************}
Unit IMMOREGFR_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     hdb,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     fe_main,
     Fiche,
     FichList,
{$else}
     MaineAGL,
     eFiche,
     eFichList,
{$ENDIF}
     {Graphics,}
     Math,
     Menus,
     HTB97,
     Variants,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     LookUp,
     AGLInit, // TheTOB
     UTob
     {Amortissements}
     ,Iment
     ,ImOutGen
     ,ImOuPlan
     ,Outils
     ,ParamSoc
      ;

Type
  TOM_IMMOREGFR = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnLoadAlerte               ; override ;
  private
    {Variables}
    fLeMode: TActionFiche;
    fOnSaveKeyDownEcran: procedure(Sender:TObject; var Key:Word; Shift:TShiftState) of object;
    fDatePassage : TDateTime;
    fChangeField : Boolean;
    fDateAchat : TDateTime;
    fMethodeEco : string;
    fVV : string;
    fDureeRestJ : integer;

    {Méthodes}
    procedure OnbValiderClick(Sender: TObject);
    procedure OnbFermeClick(Sender: TObject);
    procedure OnKeyDownEcran(Sender:TObject; var Key:Word; Shift:TShiftState);
    procedure OnAfterFormShowEcran;
    procedure OnCompteElipsisClick(Sender: TObject);
    procedure OnExitValeurVenale(Sender: TObject);
    function  ControleCode: integer;
    function  ControleCompte: integer;
    function  ControleDate: integer;
    function  ControleChamps: integer;
    function  ControleDuree: integer;
    function  ControleImmo: integer;
    procedure CalculVO_Reevaluee;
    procedure CalculDureeRestJ;
    procedure CalculDureeAmort;
    procedure CalculVNCAmort;
  end ;


////////////////////////////////////////////////////////////////////////////////
procedure AMLanceFiche_FicheREGFR (Lequel: string; LeMode: TActionFiche);
function  AMNouveauCodeImmoREGFR : string ;
procedure AMCalculREGFR (TImmoRegFR : TOB; TCalcul : TOB);
////////////////////////////////////////////////////////////////////////////////


Implementation
uses AMListeREGFR_TOF,
     AMCrREGFR_TOF;

const
{$IFDEF SERIE1}
  TAGGENERAUX=3 ;
{$ELSE}
  TAGGENERAUX=1 ;
{$ENDIF SERIE1}
  tPRO = 'PRO';

  HM: array[1..16] of string =
  {1}('Le code immobilisation est incorrect. Vous devez le modifier.'
  {2} ,'Le code que vous avez saisi existe déjà. Vous devez le modifier.'
  {3} ,'Vous devez renseigner le libellé.'
  {4} ,'Le compte d''immobilisation n''est pas renseigné.'
  {5} ,'Le compte d''immobilisation n''est pas renseigné correctement.'
  {6} ,'La date d''achat doit être antérieure strictement à la date de changement de régime.'
  {7} ,'Vous devez renseigner la méthode d''amortissement.'
  {8} ,'Vous devez renseigner la date d''achat.'
  {9} ,'Montant négatif ou nul !'
  {10} ,'Montant négatif !'
  {11} ,'Le montant des amortissements pratiqués ne peut être négatif.'
  {12} ,'La validation de l''immobilisation a échoué.'
  {13} ,'La durée d''amortissement est supérieure à 180 mois (15 ans) et le mode est dégressif.'
       + #10#13' Il convient peut-être de modifier les valeurs saisies, voulez-vous les modifier ?'
  {14} ,'La valeur retenue ne peut pas être nulle.'
  {15} ,'%d mois - %d jours'
  {16} ,'Les amortissements déjà pratiqués ne peuvent être supérieurs à la valeur retenue pour la base.'
      );


{////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////

     P R O C E D U R E S / F O N C T I O N S   P U B L I Q U E S

/////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////}

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 12/10/2007
Modifié le ... :   /  /
Description .. : Appel fenêtre de saisie d'une immo provisoire (spécifs agricoles)
Mots clefs ... :
*****************************************************************}
procedure AMLanceFiche_FicheREGFR (Lequel: string; LeMode: TActionFiche);
var
  stAction: string;
begin
  if LeMode = taCreat then
    stAction := 'ACTION=CREATIONSERIE'
  else if LeMode = taCreatOne then
    stAction := 'ACTION=CREATION'
  else if LeMode = taModif then
    stAction := 'ACTION=MODIFICATION'
  else
    stAction := 'ACTION=CONSULTATION';
  stAction := stAction + ';' + Lequel;

  AGLLanceFiche('AM', 'AMFICREGFR', '', Lequel, stAction);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 12/10/2007
Modifié le ... : 29/10/2007
Description .. : Calcul prochain code disponible (10 caractères)
Suite ........ : dans la table IMMOREGFR
Mots clefs ... :
*****************************************************************}
function AMNouveauCodeImmoREGFR : string ;
var Q : TQuery ; iCode: int64 ;
begin
  iCode:=1 ;
  Q:=OpenSql('SELECT IR_COMPTEIMMO FROM IMMOREGFR ORDER BY IR_COMPTEIMMO DESC',true) ;
  while (not Q.Eof) and (iCode=1) do
  begin
    try
      iCode:=StrToInt64(Q.Fields[0].AsString) + 1;
    except
      iCode:=1 ;
    end ;
    Q.Next ;
  end ;
  Ferme(Q) ;
  result:=Format('%.10d',[iCode]) ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 12/10/2007
Modifié le ... :   /  /
Description .. : Alimentation TOB par des calculs utilisés lors de la validation d'immo
Mots clefs ... :
*****************************************************************}
procedure AMCalculREGFR (TImmoRegFR : TOB; TCalcul : TOB);
var VO_Reevaluee, VRetenue, wFloat, VNC : double;
    wInt, wDuree : integer;
begin
 // VO réévaluée
 VO_Reevaluee := TimmoREGFR.GetValue ('IR_VO') * TimmoREGFR.GetValue ('IR_COEF');
 TCalcul.AddChampSupValeur('VO_REEVALUEE', VO_Reevaluee);

 // Durée totale amortissement
 wInt := TimmoREGFR.GetValue('IR_DUREERESTJ') div 30;
 wFloat := TimmoREGFR.GetValue('IR_DUREERESTJ') / 30;
 if wInt <> wFloat then wDuree := wInt + 1
 else wDuree := wInt;
 if TimmoREGFR.GetValue('IR_METHODEECO') = 'NAM' then
    TCalcul.AddChampSupValeur('DUREE_AMORT', 0)
 else
    TCalcul.AddChampSupValeur('DUREE_AMORT', wDuree + TImmoREGFR.GetValue ('IR_DUREERESTM'));

 // VNC théorique
 if TimmoREGFR.GetValue('IR_METHODEECO') = 'NAM' then
 begin
    VNC := VO_Reevaluee;
 end else
 begin
    if (wDuree + TImmoREGFR.GetValue ('IR_DUREERESTM')) > 0 then
      VNC := VO_Reevaluee*TImmoREGFR.GetValue('IR_DUREERESTM')/(wDuree+TImmoREGFR.GetValue ('IR_DUREERESTM'))
    else
      VNC := 0;
 end;
 TCalcul.AddChampSupValeur ('VNC', VNC);

 // Valeur retenue pour la base éco
 if TimmoREGFR.GetValue('IR_VV') = 0 then
    VRetenue := VNC
 else
    VRetenue := MinValue([VNC, TImmoREGFR.GetValue('IR_VV')]);
 TCalcul.AddChampSupValeur ('VRETENUE', VRetenue);

 // Amortissements déjà pratiqués
 if TimmoREGFR.GetValue('IR_METHODEECO') = 'NAM' then
    TCalcul.AddChampSupValeur ('AMORT', 0)
 else
    TCalcul.AddChampSupValeur ('AMORT', (TImmoREGFR.GetValue('IR_VO') - VRetenue));
end;


{////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////

       T O M   D E   L A   F O R M    I M M O R E G F R

/////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////}


procedure TOM_IMMOREGFR.OnNewRecord ;
begin
  Inherited ;
  SetField('IR_COMPTEIMMO', AMNouveauCodeImmoREGFR);
  SetField('IR_COEF', 1);
  SetControlText('VV', '');
end ;

procedure TOM_IMMOREGFR.OnDeleteRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /    
Description .. : OnUpadteRecord
Mots clefs ... : 
*****************************************************************}
procedure TOM_IMMOREGFR.OnUpdateRecord ;
var  iErreur, nJour : integer;
begin
  Inherited ;

  // Contrôles
  iErreur := ControleImmo;
  LastError := iErreur;

  if iErreur > 0 then
    LastErrorMsg := HM[iErreur]
  else if iErreur < 0 then
    LastErrorMsg := '';

  // Enregistrement
  if (iErreur = 0) then
  begin
     if GetField('IR_METHODEECO') <> 'NAM' then
     begin
       nJour := NombreJour360 (GetField ('IR_DATEACHAT'), fDatePassage-1);
       if nJour < 0 then
          nJour := 0;
       SetField ('IR_DUREERESTJ', nJour);
     end;
     SetField ('IR_VV', Valeur (GetControlText('VV')));
     if Valeur (GetControlText('VV')) > 0 then
        SetField ('IR_TOPVV', 'X')
     else
        SetField ('IR_TOPVV', '-');
  end;
end ;


procedure TOM_IMMOREGFR.OnAfterUpdateRecord ;
begin
  Inherited ;

  // => En série, l'enregistrement de la fiche + ouverture nelle se font automatiquement
  // si on a la propriété de la fiche TypeAction = taCreatEnSerie
  if (fLeMode = taCreatEnSerie) then
  begin
    TFFiche(Ecran).Typeaction :=taCreatEnSerie;
    TFFiche(Ecran).Monofiche :=False;
  end
  else
    TToolbarButton97(GetControl('BFerme')).Click;
end ;

procedure TOM_IMMOREGFR.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_IMMOREGFR.OnLoadRecord ;
begin
  Inherited ;
  fChangeField := False;
  fDateAchat := GetField('IR_DATEACHAT');
  fMethodeEco := GetField('IR_METHODEECO');
  fDureeRestJ := GetField('IR_DUREERESTJ');

  // Affichage valeur vénale
  if (GetField('IR_VV') = 0) or (GetField('IR_VV') = 0.00) then
     SetControlText('VV', '')
  else
     SetControlText('VV', StrFMontant(GetField('IR_VV'),15,V_PGI.OkDecV,'',True) );
  fVV := GetControlText('VV');

  // Durée déjà utilisée
  if (fLeMode in [taCreat..taCreatOne]) then
     SetControlVisible ('ti_Deja', False)
  else
  begin
     CalculDureeRestJ;
     SetControlEnabled('IR_DUREERESTM', (fMethodeEco <> 'NAM'));
     SetControlEnabled('VV', (fMethodeEco <> 'NAM'));
  end;

  // Autres champs calculés
  CalculVO_Reevaluee;
  CalculDureeAmort;
  CalculVNCAmort;
end ;

procedure TOM_IMMOREGFR.OnChangeField ( F: TField ) ;
var   stNomChamp, stCompte : string;
begin

 if DS.state in [dsBrowse] then
    exit;
 if fChangeField then
    exit;
 fChangeField := True;

 try
   Inherited ;
   stNomChamp := F.FieldName;

   //--------------------------------------------------------------------------
   if (stNomChamp = 'IR_COMPTEREF') then
   begin
     if (GetField(stNomChamp)<>'') then
     begin
        stCompte := ImBourreEtLess (GetField(stNomChamp), ImGeneTofb);
        if Presence('GENERAUX', 'G_GENERAL', stCompte) then
           SetField ('IR_COMPTEREF', stCompte);
     end;
   end else
   //--------------------------------------------------------------------------
   if (stNomChamp = 'IR_DATEACHAT') then
   begin
     CalculDureeRestJ;
     CalculDureeAmort;
     CalculVNCAmort;
   end else
   //--------------------------------------------------------------------------
   if (stNomChamp = 'IR_COEF') then
   begin
     CalculVO_Reevaluee;
     CalculVNCAmort;
   end else
   //--------------------------------------------------------------------------
   if (stNomChamp = 'IR_VO') then
   begin
     CalculVO_Reevaluee;
     CalculVNCAmort;
   end else
   //--------------------------------------------------------------------------
   if (stNomChamp = 'IR_METHODEECO') then
   begin
     if GetField(stNomChamp)= 'NAM' then
     begin
        SetControlEnabled('IR_DUREERESTM', False);
        SetField('IR_DUREERESTM', 0);
        SetControlEnabled('VV', False);
        SetControlText('VV', '');
     end else
     begin
        SetControlEnabled('IR_DUREERESTM', True);
        SetControlEnabled('VV', True);
     end;
     CalculDureeRestJ;
     CalculDureeAmort;
     CalculVNCAmort;
     fMethodeEco := GetField(stNomChamp);
   end else
   //--------------------------------------------------------------------------
   if (stNomChamp = 'IR_DUREERESTM') then
   begin
     CalculDureeAmort;
     CalculVNCAmort;
     { pour éviter la redondance avec le contrôle durée fait sur le OnUpdateRecord
     DureeI := ValeurI(GetControlText('Duree_Amort'));
     if not MinDureeAmortissement(GetField('IR_METHODEECO'), DureeI, stMsg) then
     begin
        PGIBox(stMsg);
        SetFocusControl('IR_DUREERESTM');
     end; }
   end;
   //--------------------------------------------------------------------------
{ PAS ICI : OnChangeField ne voit passer que les F.FieldName des champs de la Table IMMOREGFR
 => Faire un OnExit du champ
   if (stNomChamp = 'VV') then
   begin
     stVV := GetControlText('VV');
     if (stVV = '') or (stVV = '0') or (Arrondi(Valeur(stVV), 2) = 0.00) then
       SetControlText('VV', '')
     else
       SetControlText('VV', StrFMontant(Valeur(stVV),15,V_PGI.OkDecV,'',True));
     CalculAmortVNC;
   end; }

 finally
   fChangeField := False;
 end;
end ;


procedure TOM_IMMOREGFR.OnArgument ( S: String ) ;
var St : string;
begin
  Inherited ;
  fChangeField := False;

  SetControlVisible('BINSERT', True);
  SetControlVisible('BDELETE', True);

  {Paramètre en entrée}
  St := ReadTokenSt(S);
  if St = 'ACTION=MODIFICATION' then
    fLeMode := taModif
  else if St = 'ACTION=CREATION' then
    fLeMode := taCreatOne
  else if St = 'ACTION=CREATIONSERIE' then
    fLeMode := taCreatEnSerie
  else
    fLeMode := taConsult;

  { Alimentation de variables}
  fDatePassage := EncodeDate (GetParamSocSecur('SO_DATEREGREEL',''), 1, 1);

  { Evènements }
{$IFDEF EAGLCLIENT}
  THEdit(GetControl('IR_COMPTEREF')).OnElipsisClick := OnCompteElipsisClick;
{$ELSE}
  THDBEdit(GetControl('IR_COMPTEREF')).OnElipsisClick := OnCompteElipsisClick;
{$ENDIF}
  THEdit(GetControl('VV')).OnExit := OnExitValeurVenale;
  fOnSaveKeyDownEcran := Ecran.OnKeyDown;
  Ecran.OnKeyDown := OnKeyDownEcran;
  TFFiche(Ecran).OnAfterFormShow := OnAfterFormShowEcran;

  // Intercepter les boutons de sortie car le champ VV saisi n'est pas un champ de la table
  // En saisie dans VV + clic sur Mouette, on ne déclenchait pas le OnUpdateRecord
  TToolbarButton97(GetControl('BValider')).Onclick  := OnBValiderClick;
  TToolbarButton97(GetControl('BFerme')).Onclick  := OnBFermeClick;

  {Traitement des champs}
  if fLeMode in [taModif..taConsult] then
     SetControlEnabled('IR_COMPTEIMMO', False);
  SetControlEnabled('VO_REEVALUEE', False);
  //SetControlEnabled('IR_DUREERESTJ', False);
  SetControlEnabled('DUREE_AMORT', False);
  SetControlEnabled('VNC', False);
  SetControlEnabled('VRETENUE', False);
  SetControlEnabled('AMORT', False);
  SetControlVisible('BZoomAction', (fLeMode = taModif));
end ;

procedure TOM_IMMOREGFR.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_IMMOREGFR.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_IMMOREGFR.OnLoadAlerte;
begin
  Inherited ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /
Description .. : Traitement du CLIC sur la MOUETTE VERTE
Suite ........ : En effet : la fiche n'est pas validée si clic sur MOUETTE alorsqu'on est dans le champ VV
Suite ........ : On reste dans la fiche et OnUpdateRecord n'est pas fait
Suite ........ : => j'intercepte le clic pour envoyer moi-même la validation de la fiche
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMOREGFR.OnbValiderClick(Sender: TObject);
begin
  // Provoque une modif qui autorise la validation de la fiche
//  IF (fLeMode <> taConsult) then
//     DS.Edit;
  // Opération de validation -> OnUpdateRecord
  TFFiche(Ecran).BValiderClick(Sender);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /
Description .. : Traitement du CLIC sur CROIX ou ECHAP
Suite ........ : En effet : pas de gestion de la sortie si clic alorsqu'on est dans le champ VV
Suite ........ : On sort directement sans question Voulez-vous enregistrer les modifs ?
Suite ........ : => j'intercepte le clic pour envoyer moi-même le traitement de sortie
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMOREGFR.OnbFermeClick(Sender: TObject);
begin
  // Marquer une modif de fiche sur le champ IR_VV, si l'on vient de modifier VV
  // Sinon BFermeClick fait une sortie directe sans Voulez-vous enregistrer les modifs ?
  OnExitValeurVenale (Sender);

  // Envoyer le clic sur la CROIX
  TFFiche(Ecran).BFermeClick(Sender);
end;



{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /    
Description .. : Gestion touche dans l'écran
Mots clefs ... : 
*****************************************************************}
procedure TOM_IMMOREGFR.OnKeyDownEcran(Sender: TObject; var Key: Word;  Shift: TShiftState);
begin
//  case Key of
//  else
//    fOnSaveKeyDownEcran(Sender, Key, Shift);
//  end;
    fOnSaveKeyDownEcran(Sender, Key, Shift);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /    
Description .. : Après LoadRecord
Mots clefs ... : 
*****************************************************************}
procedure TOM_IMMOREGFR.OnAfterFormShowEcran;
begin
 // Mise à jour format d'affichage des zones
 SetControlProperty('IR_VO', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
 SetControlProperty('DUREE_AMORT', 'DisplayFormat', StrfMask(0,'', True));
 SetControlProperty('VNC', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
 SetControlProperty('VRETENUE', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
 SetControlProperty('AMORT', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
 SetControlProperty('VO_REEVALUEE', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
end;



{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /    
Description .. : Affichage liste des comptes
Mots clefs ... : 
*****************************************************************}
procedure TOM_IMMOREGFR.OnCompteElipsisClick(Sender: TObject);
var
  stWhere: string;
begin
  stWhere := 'G_GENERAL<="' + VHImmo^.CpteImmoSup + '" AND G_GENERAL>="' +
              VHImmo^.CpteImmoInf + '" AND G_NATUREGENE="IMO"';
  LookupList(TControl(Sender), TraduireMemoire('Comptes'), 'GENERAUX',
           'G_GENERAL', 'G_LIBELLE', stWhere, 'G_GENERAL', True, TAGGENERAUX);
end;


{procedure TOM_IMMOREGFR.OnExitCoef(Sender: TObject);
begin
  if (GetField('IR_COEF') <= 0) or (Arrondi (GetField('IR_COEF'),2) <= 0.00) then
  begin
     PGIBox (HM[9], 'Erreur');
     SetFocusControl ('IR_COEF');
  end else
  begin
     CalculVO_Reevaluee;
     CalculAmortVNC;
  end;
end; }

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 12/10/2007
Modifié le ... :   /  /
Description .. : OnChangeField ne voit passer que les champs de la table
Suite ........ : => Traitement du champ VV
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMOREGFR.OnExitValeurVenale(Sender: TObject);
var stVV : string;
begin
  stVV := GetControlText('VV');

  if fVV <> stVV then
  begin
    if (stVV = '') or (stVV = '0') or (Arrondi(Valeur(stVV), 2) = 0.00) then
       SetControlText('VV', '')
    else
       SetControlText('VV', StrFMontant(Valeur(stVV),15,V_PGI.OkDecV,'',True));

    CalculVNCAmort;
    fVV := GetControlText('VV');
    // Pour marquer la modif de fiche => sert pour OnbFermeClick
    SetField ('IR_VV', Valeur (GetControlText('VV')));
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /
Description .. : Contrôle code immo
Mots clefs ... :
*****************************************************************}
function TOM_IMMOREGFR.ControleCode: integer;
var
  stCode: string;
begin
  Result := 0;
  if fLeMode in [taCreat..taCreatOne] then
  begin
    stCode := GetField('IR_COMPTEIMMO');
    if ImExisteCarInter(stCode) then
       Result := 1
    else if ExisteSQL('SELECT IR_COMPTEIMMO FROM IMMOREGFR WHERE IR_COMPTEIMMO="' + stCode + '"')  then
       Result := 2;
  end;
  if Result <> 0 then
  begin
    SetFocusControl('IR_COMPTEIMMO');
  end;
  if (GetField('IR_LIBELLE') = '') then
  begin
    Result := 3;
    SetFocusControl('IR_LIBELLE');
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /    
Description .. : Contrôle compte immo
Mots clefs ... : 
*****************************************************************}
function TOM_IMMOREGFR.ControleCompte: integer;
begin
  Result := 0;

  if GetField('IR_COMPTEREF') = '' then
  begin
    SetFocusControl('IR_COMPTEREF');
    Result := 4
  end
  else if (not ControleCompteImmo(GetField('IR_COMPTEREF'), tPRO)) then
  begin
    SetFocusControl('IR_COMPTEREF');
    Result := 5;
  end
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /    
Description .. : Contrôle date d'achat
Mots clefs ... :
*****************************************************************}
function TOM_IMMOREGFR.ControleDate: integer;
begin
  Result := 0;

  if GetField('IR_DATEACHAT') = iDate1900 then
  begin
      Result := 8;
      SetFocusControl('IR_DATEACHAT');
  end else
  begin
     if GetField('IR_DATEACHAT') >= EncodeDate (GetParamSocSecur ('SO_DATEREGREEL',''),1,1)  then
     begin
       Result := 6;
       SetFocusControl('IR_DATEACHAT');
     end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /
Description .. : Contrôle autres champs saisis
Mots clefs ... :
*****************************************************************}
function TOM_IMMOREGFR.ControleChamps: integer;
begin
  Result := 0;
  //--------------------------------------------------------------------------
  if (GetField('IR_VO') <= 0) then
  begin
    Result := 9;
    SetFocusControl('IR_VO');
  end else
  //--------------------------------------------------------------------------
  if (GetField('IR_COEF') <= 0) or (Arrondi (GetField('IR_COEF'),2) <= 0.00) then
  begin
     Result := 9;
     SetFocusControl('IR_COEF');
  end else
  //--------------------------------------------------------------------------
  if VarIsNull (GetField('IR_METHODEECO')) or (GetField('IR_METHODEECO')='')then
  begin
      Result := 7;
      SetFocusControl('IR_METHODEECO');
  end else
  //--------------------------------------------------------------------------
  if (GetField('IR_DUREERESTM') < 0) then
  begin
      Result := 10;
      SetFocusControl('IR_DUREERESTM');
  end else
  //--------------------------------------------------------------------------
  if (GetField('IR_METHODEECO')<> 'NAM') and (Valeur(GetControlText('VRETENUE'))= 0) then
  begin
      Result := 14;
      SetFocusControl('VV');
  end else
  //--------------------------------------------------------------------------
  if (Valeur(GetControlText('AMORT')) < 0) then
  begin
      Result := 11;
  end;
  //if (Valeur(GetControlText('AMORT')) > Valeur(GetControlText('VRETENUE'))) then
  //begin
  //    Result := 16;
  //end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /
Description .. : Contrôle durée d'amortissement en liaison avec la méthode
Mots clefs ... :
*****************************************************************}
function TOM_IMMOREGFR.ControleDuree: integer;
var //DureeI : integer;
    stMsg : string;
begin
  Result := 0;

  //DureeI := ValeurI(GetControlText('Duree_Amort'));
  if not MinDureeAmortissement(GetField('IR_METHODEECO'), GetField('IR_DUREERESTM'), stMsg) then
  begin
     PGIBox(stMsg);
     SetFocusControl('IR_DUREERESTM');
     Result := -1;
  end else
  if (GetField('IR_METHODEECO') = 'DEG') and   (GetField('IR_DUREERESTM') > 180) then //(DureeI > 180) then
  begin
    if PGIAsk(HM[13]) = mrYes then
    begin
      SetFocusControl('IR_DUREERESTM');
      Result := -1;
    end;
  end;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /
Description .. : Contrôle immo
Mots clefs ... :
*****************************************************************}
function TOM_IMMOREGFR.ControleImmo: integer;
begin
   Result := ControleCode;
   if Result = 0 then
      Result := ControleCompte;
   if Result = 0 then
      Result := ControleDate;
   if Result = 0 then
      Result := ControleChamps;
   if Result = 0 then
      Result := ControleDuree;
end;




{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /
Description .. : Calcul Valeur d'origine réévaluée
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMOREGFR.CalculVO_Reevaluee;
begin
  SetControlText ('VO_REEVALUEE',
                  StrFMontant(GetField ('IR_VO') * GetField ('IR_COEF'), 15, V_PGI.OkDecV, '', True));
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /
Description .. : Calcul Durée déjà utilisée
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMOREGFR.CalculDureeRestJ;
var PremMois, PremAnnee, NbMois: Word;
    nJour: integer;
begin
//  oldNAM := False;
//  newNAM := False;
//  if (fMethodeEco = 'NAM') and (GetField ('IR_METHODEECO')<> 'NAM') then
//     oldNAM := True;
//  if (fMethodeEco <> 'NAM') and (GetField ('IR_METHODEECO')= 'NAM') then
//     newNAM := True;

  // Pour éviter un SetField = une modif d'un champ => Voulez-vous enregistrer les modifs
//  if oldNAM or newNAM  or (fDateAchat <> GetField ('IR_DATEACHAT')) then

    fDateAchat := GetField ('IR_DATEACHAT');

    SetControlVisible ('ti_Deja', (GetField ('IR_METHODEECO') <> 'NAM'));

    if GetControlVisible ('ti_Deja') then
    begin
       NOMBREMOIS(GetField ('IR_DATEACHAT'), fDatePassage, PremMois, PremAnnee,NbMois);
       nJour := NombreJour360 (GetField ('IR_DATEACHAT'), fDatePassage-1);
       if nJour < 0 then
          nJour := 0;
       SetControlCaption ('ti_Deja', Format(HM[15], [NbMois - 1, nJour]));
       fDureeRestJ := nJour;
    end else
    begin
       fDureeRestJ := 0;
    end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /
Description .. : Calcul Durée totale restant à amortir
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMOREGFR.CalculDureeAmort;
var wInt, wDuree : integer;
    wFloat : double;
begin
  wInt := fDureeRestJ div 30;
  wFloat := fDureeRestJ / 30;
  if wInt <> wFloat then wDuree := wInt + 1
  else wDuree := wInt;
  if VarIsNull (GetField('IR_DUREERESTM')) then
     SetField('IR_DUREERESTM', 0);

  if GetField ('IR_METHODEECO') = 'NAM' then
     SetControlText ('DUREE_AMORT', IntToStr(0))
  else
     SetControlText ('DUREE_AMORT', IntToStr(wDuree) + GetField ('IR_DUREERESTM'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2007
Modifié le ... :   /  /
Description .. : Calcul VNC, Valeur retenue et amortissements pratiqués jusque-là
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMOREGFR.CalculVNCAmort;
var Duree_Amort : double;
begin
 Duree_Amort := Valeur (GetControlText ('DUREE_AMORT'));
 if VarIsNull (GetField('IR_DUREERESTM')) then
    SetField('IR_DUREERESTM', 0);

 if GetField ('IR_METHODEECO') = 'NAM' then
 begin
   SetControlText ('VNC', StrFMontant(0, 15, V_PGI.OkDecV, '', True));
   SetControlText ('VRETENUE', StrFMontant(0, 15, V_PGI.OkDecV, '', True));
   SetControlText ('AMORT', StrFMontant(0, 15, V_PGI.OkDecV, '', True));
 end else
 begin
   if Duree_Amort > 0 then
     SetControlText ('VNC',
       StrFMontant( (Valeur(GetControlText('VO_REEVALUEE'))*GetField('IR_DUREERESTM')/Duree_Amort), 15, V_PGI.OkDecV, '', True))
   else
     SetControlText ('VNC', StrFMontant(0, 15, V_PGI.OkDecV, '', True));
   if (GetControlText ('VV')= '') or (Valeur(GetControlText ('VV'))= 0) then
     SetControlText ('VRETENUE',
        StrFMontant(Valeur(GetControlText ('VNC')), 15, V_PGI.OkDecV, '', True))
   else
     SetControlText ('VRETENUE',
        StrfMontant( MinValue([Valeur(GetControlText ('VNC')), Valeur(GetControlText ('VV'))]), 15, V_PGI.OkDecV, '', True));
   SetControlText ('AMORT',
      StrFMontant((GetField('IR_VO')-Valeur(GetControlText('VRETENUE'))), 15, V_PGI.OkDecV, '', True));
 end;
end;


Initialization
  registerclasses ( [ TOM_IMMOREGFR ] ) ;
end.
