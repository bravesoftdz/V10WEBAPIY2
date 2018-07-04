{-------------------------------------------------------------------------------------
  Version     |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
               28/01/02  JP   Création de l'unité
 1.40.002.001  18/05/04  JP   Ajout de la saisie de la chancellerie
 7.06.001.001  23/10/06  JP   Gestion des filtres Multi sociétés
 8.10.001.005  17/08/07  JP   Nouvelle gestion des soldes : il n'y a plus de recalcul des soldes
                              de la table TRECRITURE. Le recalcul ne concerne les soldes d'initialisation
--------------------------------------------------------------------------------------}
unit TofCalculSolde ;


interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_MAIN,
  {$ENDIF}
  StdCtrls, Controls, Classes, forms, sysutils, HCtrls, UTob, HEnt1, UTOF, HMsgBox;

procedure TRLanceFiche_CalculSolde(Dom , Fiche, Range, Lequel, Arguments : string);

type
  TOF_TOFCALCULSOLDE = class (TOF)
    procedure OnArgument (S : string ) ; override ;
    procedure OnClose                  ; override;
  protected
    DateDepart : string;
    General    : THValComboBox;
    Annee      : THValComboBox;
    ReInit     : Boolean;
    ListeAnnees: HTStringList;

    procedure InitCompteOnClick(Sender : TObject);
    procedure BValiderOnClick  (Sender : TObject);
    procedure BImprimeOnClick  (Sender : TObject);
    procedure GeneralOnChange  (Sender : TObject);
    procedure AnneeOnChange    (Sender : TObject);
    procedure InitAnnee;
    procedure MajGrid(aTob : TOB);
    procedure InitGrid;
  end;

implementation

uses
  ExtCtrls, Commun, HTB97, UProcSolde, UProcEcriture, UObjEtats,
  CPCHANCELL_TOF{18/05/04 : Pour l'appel à la fiche de chancellerie};

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_CalculSolde(Dom , Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, fiche, range, lequel, arguments );
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TOFCALCULSOLDE.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 500001;
  General := THValComboBox(GetControl('GENERAL'));
  {Gestion des filtres multi sociétés sur banquecp et dossier}
  General.Plus := FiltreBanqueCp(General.DataType, '', '');
  General.OnChange := GeneralOnChange;
  General.ItemIndex := 0;

  Annee := THValComboBox(GetControl('ANNEE'));
  InitAnnee;

  TRadioButton(GetControl('REINIT'  )).OnClick := InitCompteOnClick;
  TRadioButton(GetControl('RECALCUL')).OnClick := InitCompteOnClick;
  TRadioButton(GetControl('AFFICHE' )).OnClick := InitCompteOnClick;

  TToolBarButton97(GetControl('BVALIDER')).OnClick := BValiderOnClick;
  TToolBarButton97(GetControl('BIMPRIMER')).OnClick := BImprimeOnClick;

  ReInit := False;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TOFCALCULSOLDE.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ListeAnnees) then FreeAndNil(ListeAnnees);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TOFCALCULSOLDE.GeneralOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Devise : string;
begin
  Devise := RetDeviseCompte(General.Value);
  if Devise = '' then Devise := V_PGI.DevisePivot;
  SetControlCaption('DEV', Devise);
  AssignDrapeau(TImage(GetControl('IDEV')), Devise);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TOFCALCULSOLDE.AnneeOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
begin

  if GetControlText('ANNEE') <> '' then begin
    a := StrToInt(GetControlText('ANNEE'));
    m := 1;
    j := 1;
    DateDepart := DateToStr(EncodeDate(a, m, j));
  end
  else
    DateDepart := DateToStr(DebutAnnee(V_PGI.DateEntree));

  TRadioButton(GetControl('REINIT'  )).Caption := TraduireMemoire('Créer les soldes d''initialisation du ') + DateDepart;
  TRadioButton(GetControl('RECALCUL')).Caption := TraduireMemoire('Recalculer les soldes d''initialisation du ') + DateDepart;
  TRadioButton(GetControl('AFFICHE' )).Caption := TraduireMemoire('Afficher les soldes d''initialisation depuis le ') + DateDepart;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TOFCALCULSOLDE.InitCompteOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  B : Boolean;
begin
  B := TRadioButton(GetControl('REINIT')).Checked;
  SetControlEnabled('MONTANTOPE', B);
  SetControlEnabled('MONTANTVAL', B);
  SetControlEnabled('LBOPE'     , B);
  SetControlEnabled('LBVALEUR'  , B);
  ReInit := B;
  InitGrid;
  SetControlText('MONTANTOPE', '0');
  SetControlText('MONTANTVAL', '0');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TOFCALCULSOLDE.BValiderOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Compte     : HString;
  I          : Integer;
  OK         : Boolean;
  SoldePrecO : Double;
  SoldePrecV : Double;
  DateTravail: string;
  TOBSoldes  : TOB;
begin
  Compte := GetControlText('GENERAL');

  if TRadioButton(GetControl('AFFICHE')).Checked then begin
    TOBSoldes := GetListeSoldesInit(Compte, StrToDate(DateDepart));
    try
      MajGrid(TOBSoldes);
    finally
      if Assigned(TOBSoldes) then FreeAndNil(TOBSoldes);
    end;
    Exit;
  end;

  {Pour la réinitialisation, on ne travail que sur un compte}
  if (Compte = '') and ReInit then begin
    PGIBox(TraduireMemoire('Vous devez obligatoirement sélectionner un compte lors d''une réinitialisation.'));
    Exit;
  end;

  if TrShowMessage(Ecran.Caption, 20, '', '') = mrYes then begin
    OK := False;
    try
      {23/11/06 : On travail au matin du 01/01 maintenant
       DateTravail := DateToStr(StrToDate(DateDepart) + 1);}
      DateTravail := DateDepart;
      {On traite un compte donnée}
      if Compte <> '' then begin
        SoldePrecV  := Valeur(THNumEdit(GetControl('MONTANTVAL')).Text) ;
        SoldePrecO  := Valeur(THNumEdit(GetControl('MONTANTOPE')).Text) ;
        if ReInit then begin
          {JP 18/05/04 : afout de la saise de la cotation au premier janvier}
          if (GetControlText('DEV') <> V_PGI.DevisePivot) and
             (not ExisteSQL('SELECT H_TAUXREEL FROM CHANCELL WHERE H_DEVISE = "' + GetControlText('DEV') +
                           '" AND H_DATECOURS = "' + UsDateTime(StrToDate(DateDepart)) + '"')) then begin
            HShowMessage('0;' + Ecran.Caption + ';Veuillez saisir la cotation au premier janvier.;I;O;O;O', '', '');
            FicheChancel(GetControlText('DEV'), True, StrToDate(DateDepart), taCreat, True);
          end;

          {Création ou mise à jour de l'écriture de régularisation}
          if not CreerEcritureInit(StrToDate(DateDepart), GetControlText('DEV'), Compte, SoldePrecV, SoldePrecO) then begin
            AfficheMessageErreur(Ecran.Caption, 'Impossible de créer l''écriture de réinitialisation');
            Exit;
          end;
          {07/02/08 : FQ 10548 : pas de recalcul des soldes après saisie, mais simple affichage
          if IsNewSoldes then TOBSoldes := RecalculSoldeInit(Compte, StrToInt(GetControlText('ANNEE')))
                         else RecalculSolde(Compte, DateTravail, SoldePrecO, False, SoldePrecV);}
          if IsNewSoldes then TOBSoldes := GetListeSoldesInit(Compte, StrToDate(DateDepart))
                         else RecalculSolde(Compte, DateTravail, SoldePrecO, False, SoldePrecV);
        end
        else begin
          {Recalcul du solde}
          if IsNewSoldes then TOBSoldes := RecalculSoldeInit(Compte, StrToInt(GetControlText('ANNEE')))
                         else RecalculSolde(Compte, DateTravail, 0, True);
        end;
        OK := True;
      end
      else begin
        if IsNewSoldes then begin
          TOBSoldes := RecalculSoldeInit(Compte, StrToInt(GetControlText('ANNEE')));
          Ok := True;
        end
        else
          for I := 1 to General.Values.Count-1 do begin
            Compte     := General.Values[I];
            {Recalcul du solde}
            RecalculSolde(Compte, DateTravail, 0, True);
            OK := True;
          end;
      end;
      MajGrid(TOBSoldes);
    finally
      if Assigned(TOBSoldes) then FreeAndNil(TOBSoldes);
    end;

    {Message de résultat}
    if OK then TrShowMessage(Ecran.Caption, 6, '', '')
          else TrShowMessage(Ecran.Caption, 7, '', '');
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TOFCALCULSOLDE.InitAnnee;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
  n       : Integer;
begin
  {On récupère les millésimes correspondant aux exercices ouverts}
  ListeAnnees := GetListeMillesime;
  Annee.Items := ListeAnnees;
  Annee.Values := ListeAnnees;
  Annee.OnChange := AnneeOnChange;
  DecodeDate(V_PGI.DateEntree, a, m, j);
  n := Annee.Items.IndexOf(IntToStr(a));
  if n > -1 then Annee.ItemIndex := n
            else Annee.ItemIndex := Annee.Items.Count - 1;
  AnneeOnChange(Annee);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TOFCALCULSOLDE.MajGrid(aTob : TOB);
{---------------------------------------------------------------------------------------}
var
  FListe : THGrid;
begin
  FListe := THGrid(GetControl('FLISTE'));
  if Assigned(aTob) and Assigned(FListe) then begin
    InitGrid;
    aTob.PutGridDetail(FListe, True, False, '');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TOFCALCULSOLDE.BImprimeOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  TObjEtats.GenereEtatGrille(THGrid(GetControl('FLISTE')), TraduireMemoire('Liste des soldes d''initialisation'));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TOFCALCULSOLDE.InitGrid;
{---------------------------------------------------------------------------------------}
var
  FListe : THGrid;
begin
  FListe := THGrid(GetControl('FLISTE'));
  if Assigned(FListe) then begin
    FListe.RowCount := 2;

    FListe.ColTypes[1] := 'D';
    FListe.ColFormats[1] := ShortDateFormat;
    FListe.ColFormats[2] := StrFMask(V_PGI.OkDecV, '', True);
    FListe.ColFormats[3] := StrFMask(V_PGI.OkDecV, '', True);
    FListe.ColAligns[0] := taLeftJustify;
    FListe.ColAligns[1] := taCenter;
    FListe.ColAligns[2] := taRightJustify;
    FListe.ColAligns[3] := taRightJustify;
    FListe.ColAligns[4] := taCenter;

    FListe.Cells[0, 0] := TraduireMemoire('Compte');
    FListe.Cells[1, 0] := TraduireMemoire('Date');
    FListe.Cells[2, 0] := TraduireMemoire('Solde Val.');
    FListe.Cells[3, 0] := TraduireMemoire('Solde Opé.');
    FListe.Cells[4, 0] := TraduireMemoire('Dev.');

    FListe.Cells[0, 1] := '';
    FListe.Cells[1, 1] := '';
    FListe.Cells[2, 1] := '0.00';
    FListe.Cells[3, 1] := '0.00';
    FListe.Cells[4, 1] := '';
  end;
end;

initialization
  RegisterClasses([TOF_TOFCALCULSOLDE]);

end.

