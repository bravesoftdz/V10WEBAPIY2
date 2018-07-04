{ Unité : Source TOF de la FICHE : TRCONTROLEVALEUR
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 0.91          08/09/03  JP   Création de l'unité
 6.00.015.001  20/09/04  JP   FQ 10140 : Gestion du AllSelected
 6.00.015.001  22/09/04  JP   FQ 10138 : Mise à jour du champ E_DATEVALEUR et AllSelected en CWAS
 6.01.001.001  16/11/04  JP   FQ 10181 : Lors de la mise à jour de la table ECRITURE, la date
                              de cvaleur n'était pas corrctement mise à jour dans le cas d'un
                              glissement en nombre de jours
 6.01.001.001  17/11/04  JP   FQ 10181 bis : le décalage en jours ne se faisait que pour
                              un nombre positif : on ne pouvait revenir en arrière
 6.50.001.020  26/09/05  JP   FQ 10286 : On autorise maintenant toutes les éctritures
 6.53.001.002  26/01/06  JP   Branchement du SelectAll en eAgl à la demande de SIC
 7.00.001.008  29/05/06  JP   FQ 10203 : Interruption du traitement si la clef de valeur est null
 7.06.001.001  16/10/06  JP   Gestion des comptes courants
 7.06.001.001  17/11/06  JP   Gestion des écritures bordereaux
                              Gestion de la cohérence des dates de valeur par rapport à l'écriture de réinitialisation
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
unit TRCONTROLEVALEUR_TOF ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  StdCtrls, Controls, Classes, 
  {$IFNDEF EAGLCLIENT}
  db, Mul, FE_Main,HDB,
  {$ELSE}
  eMul, MaineAGL, uTob, 
  {$ENDIF}
  {$IFDEF TRCONF}
  uLibConfidentialite,
  {$ELSE}
  UTOF,
  {$ENDIF TRCONF}
  Forms, sysutils, HCtrls, HEnt1, HMsgBox, Commun, HTB97;


type
  {$IFDEF TRCONF}
  TOF_TRCONTROLEVALEUR = class (TOFCONF)
  {$ELSE}
  TOF_TRCONTROLEVALEUR = class (TOF)
  {$ENDIF TRCONF}
    procedure OnArgument(S : string); override;
  private
    DateValeur : string;
    NbJours    : Integer;
    CptCourant : Boolean; {16/10/06}

    procedure LancerValeur   (Sender : TObject);
    procedure ListeDbClick   (Sender : TObject);
    procedure DeviseChange   (Sender : TObject);
    procedure SlctAllClick   (Sender : TObject);
    procedure NoDossierChange(Sender : TObject);
  end ;

procedure TRLanceFiche_ControlValeur(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  TRDECALAGEJOUR_TOF, Constantes, UProcGen {LibereListe}, ExtCtrls {TImage}, UProcSolde, Math,
  UtilPgi;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_ControlValeur(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCONTROLEVALEUR.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
var
  Ch : string;
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;

  {16/10/06 : Est-on dans le cas du mul des comptes courants}
  CptCourant := ReadTokenSt(S) = 'C/C';
  
  {17/09/07 : FQ 10521 : gestion des topics}
  if CptCourant then Ecran.HelpContext := 50000149
                else Ecran.HelpContext := 50000148;

  TToolbarButton97(GetControl('BOUVRIR')).OnClick := LancerValeur;

  {26/01/06 : Branchement du AllSelect en eAgl : demande de SIC}
  SetControlVisible('BSELECTALL', True);
  TToolbarButton97(GetControl('BSELECTALL')).OnClick := SlctAllClick;

  {16/10/06 : On ne traite que les comptes courants mais pas les ICC !!}
  if CptCourant then
    Ch := 'TE_CODECIB = "' + CODECIBCOURANT + '"'
  else
    {On travaille sur les écritures non rapprochées et dont la date de valeur est inférieure
     ou égale au jour d'entrée : cela signifie que la date de valeur a été sous-évaluée : on va
     donc permettre à l'utilisateur de les décaler dans le futur}
    Ch := ' (TE_DATERAPPRO = "' + UsDateTime(iDate1900) + '" OR TE_DATERAPPRO IS NULL)' ;// FQ 10286 AND TE_NATURE = "' + na_Realise + '"';

  SetControlText('XX_WHERE', Ch);
  SetControlText('TE_DATEVALEUR', DateToStr(V_PGI.DateEntree));

  TFMul(Ecran).FListe.OnDblClick := ListeDbClick;
  THValComboBox(GetControl('TE_DEVISE')).OnChange := DeviseChange;

  {09/10/06 : On ne montre les contrôles NoDossier que si on est en multi sociétés}
  SetControlVisible('TE_NODOSSIER' , IsTresoMultiSoc);
  SetControlVisible('TTE_NODOSSIER', IsTresoMultiSoc);

  {16/10/06 : La tablette n'est pas la même pour les comptes courants}
  if CptCourant then begin
    THEdit(GetControl('TE_GENERAL')).DataType := tcp_Tous;
    THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(tcp_Tous, tcb_Courant, '');
  end
  else begin
    {Pour les comptes bancaires, simple filtres sur les dossiers du regroupement}
    THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(tcp_bancaire, '', '');
  end;

  if IsTresoMultiSoc then begin
    THValComboBox(GetControl('TE_NODOSSIER')).OnChange := NoDossierChange;
    THValComboBox(GetControl('TE_NODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCONTROLEVALEUR.NoDossierChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  s := GetControlText('TE_NODOSSIER');
  if CptCourant then THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(tcp_Tous, tcb_Courant, GetControlText('TE_NODOSSIER'))
                else THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(tcp_Bancaire, '', GetControlText('TE_NODOSSIER'));
  SetControlText('TE_GENERAL', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCONTROLEVALEUR.DeviseChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  AssignDrapeau(TImage(GetControl('IDEV')), GetControlText('TE_DEVISE'));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCONTROLEVALEUR.LancerValeur(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  S   : string;
  {$IFDEF EAGLCLIENT}
  F   : THGrid;
  {$ELSE}
  F   : THDBGrid;
  {$ENDIF}
  n   : Integer;
  Clef: string;
  Dt  : TDateTime;
  lDate : TStringList;
  Obj   : TObjDtValeur;

    {29/05/06 : FQ 10203 : on s'assure que la clef de valeur est renseignée
    {----------------------------------------------------------------------}
    function TestEcriture : Boolean;
    {----------------------------------------------------------------------}
    begin
      Result := True;
      if TFMul(Ecran).Q.FindField('TE_CLEVALEUR').AsString = '' then begin
        Result := False;
        HShowMessage('0;' + Ecran.Caption + ';Certaines clefs en valeur ne sont pas renseignées.'#13 +
                     'Veuillez lancer le module de réparation des dates de valeur dans '#13 +
                     'Transfert / Synchronisation / liste des écritures synchronisées.'#13#13 +
                     'Le traitement est interrompu.;W;O;O;O;', '', '');
      end;

    end;

    {----------------------------------------------------------------------}
    procedure DeplacerValeur;
    {----------------------------------------------------------------------}
    var
      SQL  : string;
      p    : Integer;
      dTmp : TDateTime; {FQ 10181}
      Dos  : string;
      Per  : Integer;
    begin
      Per := - 1;
      try
        {Mise à jour de la clé valeur qui est fonction de la date de valeur et récupération de la date de départ pour
         le recalcul des soldes. On part de la date la plus ancienne entre l'ancienne valeur et la nouvelle}
        if NbJours <> - 999 then begin {17/11/04 : FQ 10181 bis}
          Clef := RetourneCleEcriture(TFMul(Ecran).Q.FindField('TE_DATEVALEUR').AsDateTime + NbJours, StrToInt(Copy(TFMul(Ecran).Q.FindField('TE_CLEVALEUR').AsString, 7, 7)));
          Dt := TFMul(Ecran).Q.FindField('TE_DATEVALEUR').AsDateTime - 1;
          {16/11/04 : FQ 10181 : dans la requête de mise à jour de la table ECRITURE, on utilisait la variable
                      DateValeur qui est à jour dans le cas d'un glissement par date mais pas par jour}
          dTmp := TFMul(Ecran).Q.FindField('TE_DATEVALEUR').AsDateTime + NbJours;
          dt := Min(dTmp, dt); {17/11/04 : FQ 10181 bis : NbJours peut être négatif maintenant}
        end
        else begin
          dTmp := StrToDateTime(DateValeur); {FQ 10181}
          Clef := RetourneCleEcriture(StrToDate(DateValeur), StrToInt(Copy(TFMul(Ecran).Q.FindField('TE_CLEVALEUR').AsString, 7, 7)));
          if TFMul(Ecran).Q.FindField('TE_DATEVALEUR').AsDateTime > StrToDate(DateValeur) then
            Dt := StrToDate(DateValeur) - 1
          else
            Dt := TFMul(Ecran).Q.FindField('TE_DATEVALEUR').AsDateTime - 1;
        end;

        {15/11/06 : on fait la requête sur la clef opération qui est indexé
        SQL := S + ', TE_CLEVALEUR = "' + Clef + '" WHERE TE_CLEVALEUR = "' + TFMul(Ecran).Q.FindField('TE_CLEVALEUR').AsString +}
        SQL := S + ', TE_CLEVALEUR = "' + Clef + '" WHERE TE_CLEOPERATION = "' + TFMul(Ecran).Q.FindField('TE_CLEOPERATION').AsString +
               '" AND TE_NODOSSIER = "' + TFMul(Ecran).Q.FindField('TE_NODOSSIER').AsString + '"';
        ExecuteSQL(SQL);

        {17/11/06 : éventuel recalcul des soldes si la date change de millésime
         01/12/06 : Uniquement sur les flux réèls : cf. Mail de OG de ce jour}
        if TFMul(Ecran).Q.FindField('TE_NATURE').AsString = na_Realise then
          GereSoldeInit(TFMul(Ecran).Q.FindField('TE_DATEVALEUR').AsDateTime, StrToDateTime(DateValeur),
                        TFMul(Ecran).Q.FindField('TE_GENERAL').AsString,
                        TFMul(Ecran).Q.FindField('TE_MONTANTDEV').AsFloat, False);

        {26/09/05 : FQ 10286 : Si l'écriture n'est pas réalisée, on ne touche pas à la compta}
        if TFMul(Ecran).Q.FindField('TE_NATURE').AsString = na_Realise then begin
          if IsTresoMultiSoc then
            Dos := GetTableDossier(GetInfosFromDossier('DOS_NODOSSIER', TFMul(Ecran).Q.FindField('TE_NODOSSIER').AsString, 'DOS_NOMBASE'), 'ECRITURE')
          else
            Dos := 'ECRITURE';

          {17/11/06 : gestion des pièces de bordereaux}
          if TFMul(Ecran).Q.FindField('TE_QUALIFORIGINE').AsString = QUALIFCOMPTA then
            {Recherche de la valeur de E_PERIODE}
            Per := GetNumPeriodeTransac(TFMul(Ecran).Q.FindField('TE_NUMTRANSAC').AsString, TFMul(Ecran).Q.FindField('TE_JOURNAL').AsString);

          {22/09/04 : FQ 10140 : mise à jour du champ E_DATEVALEUR
           12/09/06 : Gestion du MultiDossier}
          SQL := 'UPDATE ' + Dos + ' SET E_DATEMODIF = "' + USDateTime(Now) + '" ';
          SQL := SQL + ', E_DATEVALEUR = "' + USDateTime(dTmp) + '" '; {FQ 10181}
          SQL := SQL + 'WHERE' +
                     ' E_JOURNAL = "'       + TFMul(Ecran).Q.FindField('TE_JOURNAL').AsString     + '" AND' +
                     ' E_EXERCICE = "'      + TFMul(Ecran).Q.FindField('TE_EXERCICE').AsString    + '" AND' +
                     ' E_DATECOMPTABLE = "' + UsDateTime(TFMul(Ecran).Q.FindField('TE_DATECOMPTABLE').AsDateTime) + '" AND' +
                     ' E_NUMEROPIECE = '    + TFMul(Ecran).Q.FindField('TE_NUMEROPIECE').AsString + ' AND'  +
                     ' E_NUMLIGNE = '       + TFMul(Ecran).Q.FindField('TE_CPNUMLIGNE').AsString  + ' AND'  +
                     ' E_NUMECHE = '        + TFMul(Ecran).Q.FindField('TE_NUMECHE').AsString     + ' AND';
          if Per > -1 then {17/11/06 : Gestion des bordereaux}
            SQL := SQL + ' E_PERIODE = '    + IntToStr(Per) + ' AND';
          SQL := SQL + ' E_QUALIFPIECE = "N"';
          ExecuteSQL(SQL);
        end;

        {On stocke les comptes traités ainsi que la date comptable la plus ancienne afin
         d'éviter un recalcul des soldes sur chaque ligne : on le fait à la fin du traitement}
        p := lDate.IndexOf(TFMul(Ecran).Q.FindField('TE_GENERAL').AsString);
        if p = - 1 then begin
          Obj := TObjDtValeur.Create;
          Obj.DateVal := TFMul(Ecran).Q.FindField('TE_DATEVALEUR').AsDateTime;
          lDate.AddObject(TFMul(Ecran).Q.FindField('TE_GENERAL').AsString, Obj);
        end else begin
          if TObjDtValeur(lDate.Objects[p]).DateVal > TFMul(Ecran).Q.FindField('TE_DATEVALEUR').AsDateTime then
            TObjDtValeur(lDate.Objects[p]).DateVal := TFMul(Ecran).Q.FindField('TE_DATEVALEUR').AsDateTime;
        end;

      except
        on E : Exception do raise;
      end;
    end;

begin
  {$IFDEF EAGLCLIENT}
  F := THGrid(TFMul(Ecran).FListe);
  {$ELSE}
  F := THDBGrid(TFMul(Ecran).FListe);
  {$ENDIF}

  {Aucune sélection, on sort}
  if (F.NbSelected = 0)
  {$IFNDEF EAGLCLIENT}
  and not F.AllSelected {20/09/04 FQ 10138}
  {$ENDIF}
  then begin
    HShowMessage('0;Mise à jour des dates de valeur; Aucun élément n''est sélectionné !;W;O;O;O;', '', '');
    Exit;
  end;

  NbJours := - 999; {17/11/04 : FQ 10181 bis}
  DateValeur := GetControlText('TE_DATEVALEUR');

  if TRLanceFiche_DecalageJours(DateValeur, NbJours) then begin
    {Demande d'incrémentation de la date de valeur}
    if NbJours <> - 999 then {17/11/04 : FQ 10181 bis}
      S := 'UPDATE TRECRITURE SET TE_DATEVALEUR = TE_DATEVALEUR + ' + IntToStr(NbJours)
    {Demande de nouvelle affectation de la date de valeur}
    else
      S := 'UPDATE TRECRITURE SET TE_DATEVALEUR = "' + UsDateTime(StrToDate(DateValeur)) + '"';
  end
  {L'utilisateur à abandonner, on sort}
  else
    Exit;

  lDate := TStringList.Create;
  try
    BeginTrans;
    try
      {$IFNDEF EAGLCLIENT}
      TFMul(Ecran).Q.First;
      if F.AllSelected then
        while not TFMul(Ecran).Q.EOF do begin
          {29/05/06 : FQ 10203 : on s'assure que la clef en valeur est bien renseignée}
          if TestEcriture then DeplacerValeur
                          else Break;
          TFMul(Ecran).Q.Next;
        end
      else
      {$ENDIF}

      for n := 0 to F.nbSelected - 1 do begin
        F.GotoLeBookmark(n);
        {$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(F.Row - 1);
        {$ENDIF}
        {29/05/06 : FQ 10203 : on s'assure que la clef en valeur est bien renseignée}
        if TestEcriture then DeplacerValeur
                        else Break;
      end;
      CommitTrans;
    except
      on E : Exception do begin
        RollBack;
        HShowMessage('0;' + Ecran.Caption + '; Traitement interrompu :'#13 + E.Message + ';E;O;O;O;', '', '');
      end;
    end;

    {20/09/04 : on fait le recalcul à la fin pour limiter le nombre de requêtes surtout avec
                le AllSlected}
    for n := 0 to lDate.Count - 1 do
      {Pour optimiser les temps de traitements, on ne recalcule que les soldes en valeur
      RecalculSolde(lDate[n], DateToStr(TObjDtValeur(lDate.Objects[n]).DateVal), 0, True);}
      RecalculSoldeValeur(lDate[n], DateToStr(TObjDtValeur(lDate.Objects[n]).DateVal - 1));
  finally
    LibereListe(lDate, True);
  end;

  {Raffraîchissement de la liste}
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCONTROLEVALEUR.ListeDbClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  {$IFDEF EAGLCLIENT}
  if TFMul(Ecran).FListe.RowCount = 0 then Exit;
  {$ELSE}
  if TFMul(Ecran).Q.Eof and TFMul(Ecran).Q.Bof then Exit;
  {$ENDIF}
  s := GetField('TE_NODOSSIER') + ';' + GetField('TE_NUMTRANSAC') + ';' +
       VarToStr(GetField('TE_NUMEROPIECE')) + ';' + VarToStr(GetField('TE_NUMLIGNE'));
  AGLLanceFiche('TR', 'TRFICECRITURE', '', s, GetField('TE_NATURE') + ';');
end;

{26/01/06 : Branchement du AllSelect en eAgl : demande de SIC
{---------------------------------------------------------------------------------------}
procedure TOF_TRCONTROLEVALEUR.SlctAllClick(Sender: TObject);
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

initialization
  RegisterClasses([TOF_TRCONTROLEVALEUR]);

end.
