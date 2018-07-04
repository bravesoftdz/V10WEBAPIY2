{-------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
               12/11/01  BT   Création de l'unité
 1.90.xxx.xxx  25/06/04  JP   Génération des écritures de trésorerie si validation
                              back office des virements FQ 10106
 6.00.015.001  20/09/04  JP   FQ 10137 : Utilisation de la propriété ForcePresentationChange au lieu
                              if not V_PGI.AutoSearch then TFStat(Ecran).ChercheClick;
 7.09.001.003  26/12/06  JP   FQ 10387 : Paramétrage de la date de virement
 8.00.001.021  20/06/07  JP   FQ 10480 : Gestion du concept VBO
--------------------------------------------------------------------------------------}
unit TofResultEqui ;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
  {$ENDIF}
  Controls, Classes, Forms, SysUtils, HCtrls, HEnt1, UTOF, HTB97, Menus,
  UTOB, Stat, ParamSoc, UObjGen;

type
  TOF_RESULTEQUI = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnClose                  ; override ;
    procedure OnArgument (S : String ) ; override ;
  protected
    PopupMenu : TPopUpMenu;
    Critere   : string;
    ManuelOk  : Boolean;
    ObjTaux   : TObjDevise;
    FDateEqui : TDateTime;

    procedure OnPopUpMenu       (Sender : TObject);
    procedure OnApresShow       ;
    procedure OnPopupNewClick   (Sender : TObject);
    procedure OnPopupModifyClick(Sender : TObject);
    procedure OnPopupDeleteClick(Sender : TObject);
    procedure bValiderOnClick   (Sender : TObject);
  end ;

function TRLanceFiche_ResultEqui(Dom, Fiche, Range, Lequel, Arguments : string) : string;

Implementation

uses
  Commun, AGLInit, TofEquilibrage, TofModifVir, HMsgBox, UProcEcriture, Constantes,
  UtilPgi, cbpPath;


{---------------------------------------------------------------------------------------}
function TRLanceFiche_ResultEqui(Dom, Fiche, Range, Lequel, Arguments : string) : string;
{---------------------------------------------------------------------------------------}
begin
  {!!!! Problème sur le retour en eAgl}
  Result := AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_RESULTEQUI.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  FreeAndNil(ObjTaux);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_RESULTEQUI.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  ManuelOk := Copy(S, 1, 6) = 'MANUEL';
  if ManuelOk then ReadTokenSt(S);
  {26/12/06 : FQ 10387 : La Date d'équilibrage est maintenant paramétrable}
  FDateEqui := V_PGI.DateEntree;
  FDateEqui := StrToDate(ReadTokenSt(S));

  {Clause where sur la nature des écritures}
  Critere := S;

  {Création de l'objet qui va servir à la conversion des soldes dans la devise d'affichage}
  ObjTaux := TObjDevise.Create(FDateEqui);

  {Paramétrage du TobViewer}
  TFStat(Ecran).LaTob := TFStat(Ecran).LaTof.LaTob;
  TFStat(Ecran).ModeAlimentation := maTOB;
  {Définition du PopupMenu de création, suppression, modification de virements}
  PopupMenu := TPopUpMenu(GetControl('POPUPMENU'));
  PopupMenu.OnPopup := OnPopupMenu;
  PopupMenu.Items[0].OnClick := OnPopupNewClick;
  PopupMenu.Items[1].OnClick := OnPopupModifyClick;
  PopupMenu.Items[2].OnClick := OnPopupDeleteClick;
  TToolBarButton97(GetControl('BVALIDER')).OnClick := bValiderOnClick;
  {On affiche toujours le résultat}
  TFStat(Ecran).OnAfterFormShow := OnApresShow;
  ADDMenuPop(PopupMenu, '', '');
  {JP 20/08/04 : FQ 10122, c'est le seul moyen que j'ai trouvé de contourné le repositionnement des boutons}
  TToolBarButton97(GetControl('BIMPRIMER1')).OnClick := TToolBarButton97(GetControl('BIMPRIMER')).OnClick;

  {20/09/04 : FQ 10137 : Affichege automatique de la présentation}
  TFStat(Ecran).ForcePresentationChange := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_RESULTEQUI.OnApresShow;
{---------------------------------------------------------------------------------------}
begin
  {En ouverture de fiche, si l'on est en mode manuel, on ouvre l'écran de création d'un
   nouveau virement pour éviter de laisser le client perplexe devant un écran vide}
  if ManuelOk then
    OnPopupNewClick(PopupMenu);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_RESULTEQUI.OnPopupMenu(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  OK: Boolean;
begin
  OK := TFStat(Ecran).TV.CurrentTOB <> Nil; // Pas de Tob courante -> New seul possible
  PopupMenu.Items[1].Enabled := OK; // Modify
  PopupMenu.Items[2].Enabled := OK; // Delete
end;

{---------------------------------------------------------------------------------------}
procedure TOF_RESULTEQUI.OnPopupNewClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  TheData := TFStat(Ecran).LaTob;
  {26/12/06 : FQ 10387 : Ajout de la date d'équilibrage}
  if TRLanceFiche_ModifVir('TR','TRMODIFVIR', '', '', 'CRE;' + DateToStr(FDateEqui) + ';' + Critere) <> '' then
    {Rafraichissement de l'écran puisque la Tob a été modifiée}
    TFStat(Ecran).ChercheClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_RESULTEQUI.OnPopupModifyClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {11/08/04 : pour éviter un violation d'accès si on n'est pas sur une ligne de détail}
  if TFStat(Ecran).TV.CurrentTOB = nil then Exit;
  TheData := TFStat(Ecran).TV.CurrentTOB;
  {On ne passe que la devise affichée
   26/12/06 : FQ 10387 : Ajout de la date d'équilibrage}
  TRLanceFiche_ModifVir('TR','TRMODIFVIR', '', '', 'MOD;' + DateToStr(FDateEqui) + ';' + Critere); {22/06/04 : problème d'agl}
  {Rafraichissement de l'écran puisque la Tob a été modifiée}
  TFStat(Ecran).ChercheClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_RESULTEQUI.OnPopupDeleteClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  aTob    : TOB;
  Mnt, MD,
  SSolde,
  DSolde  : Double;
  SCompte,
  DCompte : string;
begin
  {Suppression de la ligne}
  aTob := TFStat(Ecran).TV.CurrentTOB;
  Mnt := aTob.GetValue('MONTANT');
  SCompte := aTob.GetValue('SCOMPTE');
  DCompte := aTob.GetValue('DCOMPTE');
  SSolde := aTob.GetValue('SSOLDE') + Mnt; // Retire le virement
  ObjTaux.ConvertitMnt(Mnt, MD, ObjTaux.GetDeviseCpt(SCompte), DCompte);
  DSolde := aTob.GetValue('DSOLDE') - MD;
  aTob.Free;
  {Mise à jour des soldes}
  TOF_EQUILIBRAGE.MajSolde(TFStat(Ecran).LaTob, SCompte, SSolde); // ou aTob.Parent
  TOF_EQUILIBRAGE.MajSolde(TFStat(Ecran).LaTob, DCompte, DSolde);
  {Rafraichissement du TobViewer}
  TFStat(Ecran).ChercheClick;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 31/12/2001
Modifié le ... :   /  /
Description .. : Ecrit dans la table EQUILIBRAGE les virements à effectuer
Mots clefs ... : VIREMENT;EQUILIBRAGE
*****************************************************************}
procedure TOF_RESULTEQUI.bValiderOnClick(Sender : TObject);
var
  Q        : TQuery;
  I, N, Err: Integer;
  SCompte,
  Banque,
  SClef, DClef,
  RacineNom,
  Devise,
  Sql      : string;
  Mnt,
  Cotation : Double;
  aTob     : TOB;
  OkVBO    : Boolean;
  TM, TF   : TOB;
begin
  {28/06/04 : Pour simplifier l'utilisation de CreeTrEcritureFromVirement, j'ai supprimé la création des
              virements par requêtes et j'utilise une tob (TM).
              Par ailleurs, s'il y a validation Back-Office, le traitement s'effectue en trois temps :
              1/ Création des virements TM.InsertDB
              2/ Création des TREcritures CreeTrEcritureFromVirement
              3/ Mise à jour des virements TM.UpdateDB
              Certes on effectue 2 requêtes sur la Table équilibrage mais la création des virements est
              indépendante de celle des écritures et donc il n'y a pas de raison de créer les écritures
              puis les virements le tout encapsulé par une transaction (d'ailleurs il y a une transaction
              dans CreeTrEcritureFromVirement)}

  {25/06/04 : Gestion des erreurs lors de la création des écritures de trésorerie}
  Err := 0;

  {31/03/05 : FQ 10223 : Nouvelle gestion des erreurs : initialisation des variables globales}
  InitGestionErreur(CatErr_TRE);

  if (TFStat(Ecran).LaTob.Detail.Count > 0) and (trShowMessage(Ecran.Caption, 20, '', '') = mrYes) then begin
    {Recherche du dernier numéro utilisé}
    Q := OpenSql('SELECT MAX(TEQ_NUMEQUI) FROM EQUILIBRAGE', True);
    if Q.Eof then N := 0
             else N := Q.Fields[0].AsInteger;
    Ferme(Q);

    RacineNom := GetParamSocSecur('SO_CPRACINEVIREMENT', TcbpPath.GetCegidUserLocalAppData);

    {La validation peut être manuelle (dans la fiche virements) ou automatique (ici)
     20/06/07 : FQ 10480 : gestion du concept VBO}
    OkVBO := GetParamSocSecur('SO_VIREMENTVBO', False) and CanValidateBO;

    BeginTrans;
    try
      TM := TOB.Create('£EQUILIBRAGE', nil, -1);
      try
        for I := 0 to TFStat(Ecran).LaTob.Detail.Count - 1 do begin
          TFStat(Ecran).Retour := 'X';
          Inc(N);
          aTob := TFStat(Ecran).LaTob.Detail[I];
          SCompte := aTob.GetValue('SCOMPTE');

          {Recherche de la banque du compte}
          Q := OpenSQL('SELECT BQ_BANQUE FROM BANQUECP WHERE BQ_CODE = "' + SCompte + '"', True);
          if not Q.EOF then
            Banque := Q.Fields[0].AsString
          else begin
            {La banque n'a pas été trouvée ...}
            MessageAlerte('Impossible de trouver la banque du compte : ' + SCompte);
            {... on passe au compte suivant}
            Continue;
          end;
          Ferme(Q);

          {La devise est passée dans les trois premiers caractères des critères}
          Devise := aTob.GetValue('SDEVISE');
          Mnt := aTob.GetValue('MONTANT');
          {Récupération de la cotation de la devise par rapport à la date du jour ??}
          Cotation := ObjTaux.GetParite(V_PGI.DevisePivot, Devise);
          if Cotation = 0 then Cotation := 1;

          TF := TOB.Create('EQUILIBRAGE', TM, -1);
          TF.SetInteger('TEQ_NUMEQUI', N);
          {26/12/06 : FQ 10387 : Paramétrage de la date du virement}
          TF.SetDateTime('TEQ_DATECREATION', aTob.GetDateTime('DATEEQUI'));
          TF.SetString('TEQ_USERCREATION', V_PGI.User);
          if IsTresoMultiSoc then begin
            TF.SetString('TEQ_SNODOSSIER', GetDossierFromBQCP(SCompte));
            TF.SetString('TEQ_DNODOSSIER', GetDossierFromBQCP(aTob.GetString('DCOMPTE')));
            TF.SetString('TEQ_SOCIETE', GetInfosFromDossier('DOS_NODOSSIER', TF.GetString('TEQ_SNODOSSIER'), 'DOS_SOCIETE'));
          end else begin
            TF.SetString('TEQ_SOCIETE', V_PGI.CodeSociete);
            TF.SetString('TEQ_SNODOSSIER', V_PGI.NoDossier); {29/05/06 : FQ 10360}
            TF.SetString('TEQ_DNODOSSIER', V_PGI.NoDossier); {29/05/06 : FQ 10360}
          end;
          TF.SetString('TEQ_SBANQUE', Banque);
          TF.SetString('TEQ_SGENERAL', SCompte);
          TF.SetString('TEQ_DGENERAL', aTob.GetString('DCOMPTE'));
          TF.SetString('TEQ_DEVISE', Devise);
          TF.SetString('TEQ_FICVIR', RacineNom + Banque + TF.GetString('TEQ_SOCIETE') + Devise + FormatDateTime('ddmmyyhhnn', Now)  + '.txt');
          TF.SetString('TEQ_IMPRIME', '-');
          TF.SetString('TEQ_FICEXPORT', '-');
          TF.SetString('TEQ_CLESOURCE', '');
          TF.SetString('TEQ_CLEDESTINATION', '');
          TF.SetDouble('TEQ_COTATION', Cotation);
          TF.SetDouble('TEQ_MONTANTDEV', Mnt);
          TF.SetDouble('TEQ_MONTANT', Mnt / Cotation);
          TF.SetString('TEQ_VALIDBO', '-');
        end;

        {Insertion des virements}
        //TM.InsertDB(nil, True);

        {25/06/04 : Génération des écritures de trésorerie si validation BO (FQ 10106)}
        if OkVBO then begin
          for i := 0 to TM.Detail.Count - 1 do begin
            TF := TM.Detail[i];
            {Création des deux écritures correspondant au virement courant}
            if not CreeTrEcritureFromVirement(TF, ObjTaux, False, SClef, DClef) then begin
              Inc(Err);
            end
            else begin
              {Si la création des écritures s'est normalement déroulée, mise à jour des virements}
              TF.SetString('TEQ_CLESOURCE', SClef);
              TF.SetString('TEQ_CLEDESTINATION', DClef);
              TF.SetString('TEQ_VALIDBO', 'X');
            end;
          end;
          {Mise à jour des virements}
          //TM.UpdateDB(True);
          {Gestion des éventuelles erreurs}
          if Err > 0 then begin
            if Err = 1 then SQL := IntToStr(Err) + ' virement n''a pas été validé back-office : ' {16/03/05 : FQ 10223}
                       else SQL := IntToStr(Err) + ' virements n''ont pas été validés back-office :';
            AfficheMessageErreur(Ecran.Caption, SQL);
          end;
        end;

        {Insertion des virements}
        TM.InsertDB(nil, True);
        CommitTrans;
      finally
        FreeAndNil(TM);
      end;
    except
      on E : Exception do begin
        RollBack;
        PGIError(TraduireMemoire('Impossible de créer les virements : ') + #13#13 + E.Message);
      end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_RESULTEQUI.OnNew ;
{---------------------------------------------------------------------------------------}
begin
  Inherited;
  {Mettre nil semblerait produire des fuites de mémoire en mode MemCheck}
  OnPopupModifyClick(Self);
end ;

Initialization
  registerclasses ( [ TOF_RESULTEQUI ] ) ;
end.

