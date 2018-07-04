{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 04/02/2003
Modifié le ... : 11/02/2004
Description .. : - Passage en eAGL
Suite ........ : - FQ 10408 - CA - 22/04/2003
Suite ........ : - FQ 10597 - CA - 22/04/2003 - Libellé exercice.
Suite ........ : - FQ 12259 - CA - 22/04/2003 - Contrôle durée d'exercice
Suite ........ : - 26/08/2003 - CA - FQ 12644 - Initialisation des champs exercice
Mots clefs ... :
*****************************************************************}
unit OuvreExo;

interface

uses Windows, Classes, Controls, StdCtrls, ExtCtrls, Mask,
  Messages, // WM_CLOSE
  SysUtils, // DecodeDate, EncodeDate, FormatFloat, StrToInt, DateToStr, StrToDate, FormatDateTime
  Forms, // TForm, TCloseAction, Screen, ShowModal, Show, Handle, caFree
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}

{$ENDIF}

  Ent1, // TExoDate, ChargeMagExo, VH, NombrePerExo
  HEnt1, // String3, SyncrDefault, V_PGI, FinDeMois, IsValidDate
  HPanel, // THPanel
  UiUtil, // FindInsidePanel, InitInside
  UTOB,
  {$IFDEF NETEXPERT}
  UNEACtions,
  {$ENDIF}
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  hmsgbox, HSysMenu, Hctrls, HTB97;

procedure OuvertureExo;
function PrepareOuvreExoSansCompta: boolean;
function TrouveNewCodeExo: string;
function LibelleExerciceDefaut(DateDebut, DateFin: TDateTime; bAbrege: boolean = False): string;

// FQ 13355 SBO 08/08/2005 pour utilisation dans Ent1...
function TestExistExoNPlus1(ExoEnCours: TExoDate): string;
function GenDatesExoN(var ExoEncours: TExoDate): Boolean;
function JeLePeux: boolean;
procedure OuvertureExoModal;

type
  TOuvExo = class(TForm)
    Panel1: TPanel;
    Confirmation: THMsgBox;
    HLabel1: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    FDateEnCDeb: TMaskEdit;
    FDateEnCFin: TMaskEdit;
    FDateSuivDeb: TMaskEdit;
    FDateSuivFin: TMaskEdit;
    Morceaux: THMsgBox;
    HMTrad: THSystemMenu;
    Dock971: TDock97;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FDateSuivFinChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BFermeClick(Sender: TObject); // FQ 12422
  private
    { Déclarations privées }
    ExoEnCours, ExoSuivant: TExoDate;
    CodOuvert: String3;
    EtatExo: Integer;
    function ExistExoNPlus1: Boolean;
    function OKDateFinNPlus1: Boolean;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  UtilPgi,
  Cloture ,
  ParamSoc,
  UlibRevision,
  uLibExercice; // CControleDureeExercice

function PrepareOuvreExoSansCompta: boolean;
var
  Q: TQuery;
  TOBX: TOB;
  St, CodeExercice: string;
  Okok, ExMvt, FermerPrec: Boolean;
begin
  FermerPrec := False;
  TOBX := TOB.Create('LESEXERCICES', nil, -1);
  Q :=
    OpenSQL('SELECT * FROM EXERCICE WHERE EX_ETATCPTA="OUV" ORDER BY EX_DATEDEBUT', True);
  TOBX.LoadDetailDB('EXERCICE', '', '', Q, False);
  Ferme(Q);
  if TOBX.Detail.Count = 0 then
  begin
    okOk := True;
    Q := OpenSQL('SELECT * FROM EXERCICE ORDER BY EX_DATEDEBUT DESC', True);
    if not Q.EOF then
    begin
      CodeExercice := Q.Findfield('EX_EXERCICE').asString;
      if not ExisteSQL('SELECT E_JOURNAL FROM ECRITURE WHERE E_EXERCICE="' +
        CodeExercice + '"') then // c'est qu'il vient d'être créé
      begin
        ExecuteSQL('UPDATE EXERCICE SET EX_ETATCPTA="OUV" WHERE EX_EXERCICE="' +
          CodeExercice + '"');
        OkOk := False;
        PgiInfo('Exercice ouvert', 'Ouverture d''exercice');
      end;
    end;
    Ferme(Q);
  end
  else if TOBX.Detail.Count < 2 then
  begin
    Okok := True;
  end
  else
  begin
    ExMvt := ExisteSQL('SELECT E_JOURNAL FROM ECRITURE WHERE E_EXERCICE="' +
      TOBX.Detail[0].GetValue('EX_EXERCICE') + '"');
    if not ExMvt then
    begin
      FermerPrec := True;
      Okok := True;
    end
    else
    begin
      Okok := False;
      St := 'Vous avez 2 exercices ouverts et mouvementés. ' + #10#13;
      St := St + 'Vous devez au préalable effectuer une clôture comptable. ' +
        #10#13;
      St := St + 'Si vous n''utilisez pas la comptabilité, vous pouvez forcer '
        + #10#13;
      St := St + 'l''ouverture d''un nouvel exercice en cliquant sur Ok ' +
        #10#13;
      St := St + 'sinon cliquez sur Annuler ';
      if HShowMessage('5;Ouverture d''exercice;' + St + ';Q;OC;C;C', '', '') =
        mrOk then
      begin
        FermerPrec := True;
        Okok := True;
      end;
    end;
    if FermerPrec then
    begin
      TOBX.Detail[0].PutValue('EX_ETATCPTA', 'CDE');
      TOBX.UpdateDB;
      ChargemagExo(False);
    end;
  end;
  TOBX.Free;
  Result := Okok;
end;

function GenDatesExoN(var ExoEncours: TExoDate): Boolean;
var
  Q: TQuery;
begin
  if VH^.EnCours.Code = '' then
  begin
    Result := False;
    Exit;
  end;
  Q := OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN, EX_ETATCPTA FROM EXERCICE' +
    ' WHERE EX_EXERCICE="' + VH^.EnCours.Code + '"', TRUE);
  if not Q.EOF then
  begin
    ExoEnCours.Code := VH^.EnCours.Code;
    ExoEnCours.Deb := Q.FindField('EX_DATEDEBUT').asDateTime;
    ExoEnCours.Fin := Q.FindField('EX_DATEFIN').asDateTime;
    Result := True;
  end
  else
    Result := FALSE;
  Ferme(Q);
  Screen.Cursor := SyncrDefault;
end;

function TestExistExoNPlus1(ExoEnCours: TExoDate): string;
var
  Q: TQuery;
  DD: tDateTime;
begin
  Result := '';
  DD := ExoEnCours.Fin + 1;
  Q := OpenSQL('SELECT EX_ETATCPTA FROM EXERCICE' + ' WHERE EX_DATEDEBUT="' +
    UsDateTime(DD) + '"', TRUE);
  if not Q.EOF then
  begin
    Result := Q.FindField('EX_ETATCPTA').asString;
  end;
  Ferme(Q);
end;

function JeLePeux: boolean;
var
  ad, md, NbMois: Word;
  ExoEnCours: TExoDate;
begin
  Result := False;
  NombrePerExo(VH^.EnCours, md, ad, NbMois);
  if not GenDatesExoN(ExoEnCours) then
  begin
    HShowMessage('4;Ouverture d''exercice;L''exercice en cours n''est pas ouvert ou n''existe pas. Vous devez le paramétrer par l''assistant de paramétrage de société.;W;O;O;O', '', '');
    Exit;
  end;
  if TestExistExoNPlus1(ExoEnCours) = 'OUV' then
  begin
    HShowMessage('0;Ouverture d''exercice;L''exercice suivant est déjà ouvert.;E;O;O;O', '', '');
    Exit;
  end;
  Result := True;
end;

procedure OuvertureExo;
var
  OuvExo: TOuvExo;
  PP: THPanel;
begin
  if not _BlocageMonoPoste(True) then
    Exit;
  if V_PGI.OutLook then
    if not JeLePeux then
      Exit;
  OuvExo := TOuvExo.Create(Application);
  PP := FindInsidePanel;
  if PP = nil then
  begin
    try
      OuvExo.ShowModal;
    finally
      OuvExo.Free;
      _DeblocageMonoPoste(True);
    end;
    Screen.Cursor := SyncrDefault;
  end
  else
  begin
    InitInside(OuvExo, PP);
    OuvExo.Show;
  end;
end;

procedure OuvertureExoModal;
var OuvExo : TOuvExo;
begin
  if not _BlocageMonoPoste(True) then Exit;
  OuvExo := TOuvExo.Create(Application);

  try
    OuvExo.ShowModal;
    finally
      OuvExo.Free;
      _DeblocageMonoPoste(True);
    end;

  Screen.Cursor := SyncrDefault;

end;

function DonneAnnee(LaDate: TDateTime): Integer;
var
  a, m, j: Word;
begin
  DecodeDate(LaDate, a, m, j);
  Result := a;
end;

function DonneMois(LaDate: TDateTime): Integer;
var
  a, m, j: Word;
begin
  DecodeDate(LaDate, a, m, j);
  Result := m;
end;

procedure TOuvExo.FormShow(Sender: TObject);
var
  ad, md, NbMois: Word;
  An, Mois, Jour: Word;
  Reponse: Integer;
begin
  BValider.Enabled := True;
  NombrePerExo(VH^.EnCours, md, ad, NbMois);
  if not GenDatesExoN(ExoEnCours) then
    if not V_PGI.OutLook then
    begin
      Confirmation.Execute(4, '', '');
      PostMessage(Handle, WM_CLOSE, 0, 0);
    end;
  if ExistExoNPlus1 then
  begin
    if CodOuvert = 'OUV' then
    begin
      if not V_PGI.OutLook then
      begin
        Confirmation.Execute(0, '', '');
        PostMessage(Handle, WM_CLOSE, 0, 0);
      end;
    end
    else
    begin
      Reponse := Confirmation.Execute(1, '', '');
      EtatExo := 1;
      BValider.Enabled := False;
      if Reponse = mrYes then
      begin
        ExecuteSQL('UPDATE EXERCICE SET EX_ETATCPTA="OUV",EX_ETATBUDGET="OUV",EX_NATEXO="" WHERE EX_EXERCICE="'
          + ExoSuivant.Code + '" ');
        BValider.Enabled := False;
        // ---- On averti le serveur de la mise à jour des paramSoc et de la table exercice  ----
        AvertirCacheServer( 'EXERCICE' ) ; // SBO : 15/07/2004
        VH^.Suivant.Code := ExoSuivant.Code;
        VH^.Suivant.Deb := ExoSuivant.Deb;
        VH^.Suivant.Fin := ExoSuivant.Fin;
        // *ANODYNA
        if GetParamSocSecur('SO_CPANODYNA',false) then
         begin
          PGIInfo('Vous avez activé les A-nouveaux dynamiques, vous devez faire une cloture provisoire') ;
  {$IFNDEF ENTREPRISE}
          ClotureComptable(false) ;
  {$ENDIF}

         end ;
      end;
    end;
  end
  else
  begin
    ExoSuivant.Code := (FormatFloat('000', StrToInt(ExoEnCours.Code) + 1));
    ExoSuivant.Deb := ExoEnCours.Fin + 1;
    DecodeDate(ExoSuivant.Deb, An, Mois, Jour);
    if Mois + NbMois - 1 <= 12 then
      ExoSuivant.Fin := EncodeDate(An, Mois + NbMois - 1, 25)
    else
    begin
      if (Mois + NbMois - 1) - 12 > 12 then
        ExoSuivant.Fin := EncodeDate(An + 1, 12, 25)
      else
        ExoSuivant.Fin := EncodeDate(An + 1, (Mois + NbMois - 1) - 12, 25);
    end;
    ExoSuivant.Fin := FinDeMois(ExoSuivant.Fin);
    EtatExo := 2;
  end;
  FDateEnCDeb.Text := DateToStr(ExoEnCours.Deb);
  FDateEnCFin.Text := DateToStr(ExoEnCours.Fin);
  FDateSuivDeb.Text := DateToStr(ExoSuivant.Deb);
  FDateSuivFin.Text := DateToStr(ExoSuivant.Fin);
end;

procedure TOuvExo.FDateSuivFinChange(Sender: TObject);
begin
  if (FDateSuivFin.Text <> '  /  /    ') and IsValidDate(FDateSuivFin.Text) then
    ExoSuivant.Fin := StrToDate(FDateSuivFin.Text);
end;

function TOuvExo.ExistExoNPlus1: Boolean;
var
  Q: TQuery;
  DD: tDateTime;
begin
  Result := False;
  (*
  Q:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN, EX_ETATCPTA FROM EXERCICE'+
             ' WHERE EX_EXERCICE="'+(FormatFloat('000',StrToInt(ExoEnCours.Code)+1))+'"' ,TRUE) ;
  *)
  DD := ExoEnCours.Fin + 1;
  Q :=
    OpenSQL('SELECT EX_EXERCICE,EX_DATEDEBUT, EX_DATEFIN, EX_ETATCPTA FROM EXERCICE'
    +
    ' WHERE EX_DATEDEBUT="' + UsDateTime(DD) + '"', TRUE);
  if not Q.EOF then
  begin
    { l'exercice suivant n'a pas obligatoirement Code + 1 }
  //   ExoSuivant.Code:=(FormatFloat('000',StrToInt(ExoEnCours.Code)+1)) ;
  // CA - 22/04/2003
    ExoSuivant.Code := Q.FindField('EX_EXERCICE').AsString;
    ExoSuivant.Deb := Q.FindField('EX_DATEDEBUT').asDateTime;
    ExoSuivant.Fin := Q.FindField('EX_DATEFIN').asDateTime;
    CodOuvert := Q.FindField('EX_ETATCPTA').asString;
    FDateSuivFin.Enabled := False;
    Result := True;
  end
  else
    FDateSuivFin.Enabled := True;
  Ferme(Q);
  Screen.Cursor := SyncrDefault;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : ../../....
Modifié le ... : 11/02/2004
Description .. : GCO - 11/02/2004
Suite ........ : -> Suppression pour création d'exercice
Suite ........ : à n'importe quelles dates
Mots clefs ... :
*****************************************************************}

function TOuvExo.OKDateFinNPlus1: Boolean;
begin
  Result := CControleDureeExercice(ExoSuivant.Deb, ExoSuivant.Fin);
end;

function TrouveNewCodeExo: string;
var
  Q: TQuery;
  St: string;
  i: Integer;
begin
  Result := '010';
  Q := OpenSQL('SELECT MAX(EX_EXERCICE) FROM EXERCICE ', TRUE);
  if not Q.Eof then
  begin
    St := Q.Fields[0].AsString;
    i := StrToInt(St);
    Result := FormatFloat('000', i + 1);
  end;
  Ferme(Q);
end;

procedure TOuvExo.BValiderClick(Sender: TObject);
var
  Reponse: Integer;
  SQL, NExo: string;
  {$IFDEF NETEXPERT}
  IsNetEXpert : Boolean;
  {$ENDIF}

begin
  { FQ 20754 BVE 21.06.07 }
  if not(IsValidDate(FDateSuivDeb.Text)) then
  begin
     PGIError('La date de début d''exercice est incorrecte.');
     FDateSuivDeb.SetFocus;
     Exit;
  end
  else if not(IsValidDate(FDateSuivFin.Text)) then
  begin       
     PGIError('La date de fin d''exercice est incorrecte.');
     FDateSuivFin.SetFocus;
     Exit;
  end;
  { END FQ 20754 }

    {$IFDEF NETEXPERT}
    if not InterrogeUserService (IsNetEXpert) then
    begin
        if IsNetExpert then
        begin
            PGIInfo('Il y a des utilisateurs connectés sur le dossier Business Line, cette fonction n''est plus disponible',Caption);
            BValider.Visible := false;
            ModalResult := mrCancel;
            Exit;
        end;
    end;
    {$ENDIF}


  if OKDateFinNPlus1 then
  begin
    Reponse := Confirmation.Execute(1, '', '');
    if Reponse = mrYes then
    begin
      {T:=OpenSQL('SELECT * FROM EXERCICE WHERE EX_EXERCICE="Ede"',FALSE) ;
      T.Insert ; InitNew(T) ;
      T.FindField('EX_EXERCICE').AsString:=TrouveNewCodeExo ;
      T.FindField('EX_LIBELLE').AsString:=Morceaux.Mess[0]+DateToStr(ExoSuivant.Deb)+
                                          Morceaux.Mess[1]+DateToStr(ExoSuivant.Fin) ;
      T.FindField('EX_ABREGE').AsString:=Morceaux.Mess[2]+FormatDateTime('dd/mm',ExoSuivant.Deb) ;
      T.FindField('EX_DATEDEBUT').AsDateTime:=ExoSuivant.Deb ;
      T.FindField('EX_DATEFIN').AsDateTime:=ExoSuivant.Fin ;
      T.FindField('EX_ETATCPTA').AsString:='OUV';
      T.FindField('EX_ETATBUDGET').AsString:='OUV';
      T.FindField('EX_NATEXO').AsString:='' ;
      T.Post ;
      Ferme(T) ; }
      NExo := TrouveNewCodeExo;
      SQL := 'INSERT INTO EXERCICE';
      SQL := SQL + '(EX_EXERCICE, EX_LIBELLE, EX_ABREGE,';
      SQL := SQL + ' EX_DATEDEBUT, EX_DATEFIN, EX_ETATCPTA,';
      SQL := SQL +
        ' EX_ETATBUDGET, EX_NATEXO,EX_ETATADV,EX_ETATAPPRO,EX_ETATPROD,EX_SOCIETE,EX_VALIDEE,EX_DATECUM,EX_DATECUMRUB,EX_DATECUMBUD,EX_DATECUMBUDGET,EX_BUDJAL)';
      SQL := SQL + ' VALUES ("';
      SQL := SQL + NExo + '","';
      //      SQL := SQL + Morceaux.Mess[0]+DateToStr(ExoSuivant.Deb)+
      //                   Morceaux.Mess[1]+DateToStr(ExoSuivant.Fin) + '","' ;;
      //      SQL := SQL + Morceaux.Mess[2]+FormatDateTime('dd/mm',ExoSuivant.Deb) + '","' ;;*
      // CA - 22/04/2003 - 10597 - Libellé exercice.
      SQL := SQL + LibelleExerciceDefaut(ExoSuivant.Deb, ExoSuivant.Fin, False) +
        '","';
      SQL := SQL + LibelleExerciceDefaut(ExoSuivant.Deb, ExoSuivant.Fin, True) +
        '","';
      SQL := SQL + usdatetime(ExoSuivant.Deb) + '","';
      SQL := SQL + usdatetime(ExoSuivant.Fin) + '","';
      SQL := SQL + 'OUV' + '","';
      SQL := SQL + 'OUV' + '",';
      SQL := SQL + '"","';
      SQL := SQL + 'NON' + '","';
      SQL := SQL + 'NON' + '","';
      SQL := SQL + 'NON' + '","';
      SQL := SQL + V_PGI.CodeSociete + '","';
      SQL := SQL + '------------------------' + '","';
      SQL := SQL + UsDateTime(2) + '","';
      SQL := SQL + UsDateTime(2) + '","';
      SQL := SQL + UsDateTime(2) + '","';
      SQL := SQL + UsDateTime(2) + '",';
      SQL := SQL + '"")';
      ExecuteSQL(SQL);
      FDateSuivFin.Enabled := False;
      // ---- On averti le serveur de la mise à jour des paramSoc et de la table exercice  ----
      AvertirCacheServer( 'EXERCICE' ) ; // SBO : 15/07/2004
{$IFDEF COMPTA}
      MAJHistoparam ('EXERCICE',  NExo);
{$ENDIF}
    end;
    BValider.Enabled := False;
    VH^.Suivant.Code := ExoSuivant.Code;
    VH^.Suivant.Deb := ExoSuivant.Deb;
    VH^.Suivant.Fin := ExoSuivant.Fin;
{$IFDEF NETEXPERT}
       // Synchronisation avec BL
    if IsNetExpert then
          SynchroOutBL ('O', FDateSuivDeb.Text, FDateSuivFin.Text);
{$ENDIF}

    // GCO - 11/10/2007 - FQ 21633
    if VH^.Revision.Plan <> '' then
      SynchroRICAvecExercice;

    Confirmation.Execute(5, '', '');
  end
  else
  begin
    // GCO - 11/05/2006 - FQ 13293
    PGIInfo('La durée d''exercice ne doit pas excéder 24 mois.', 'Ouverture d''exercice');
    FDateSuivFin.SetFocus;
  end;
end;

procedure TOuvExo.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TOuvExo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
{$IFDEF EAGLCLIENT}
  AvertirCacheServer( 'EXERCICE' ) ; // GCO - 19/09/2007 - FQ 21304
{$ENDIF}
  AvertirMultiTable('TTEXERCICE');
  ChargeMagExo(False);
  if IsInside(Self) then
  begin
    _DeblocageMonoPoste(True);
    Action := caFree;
  end;
end;

procedure TOuvExo.FormCreate(Sender: TObject);
begin
  PopUpMenu := ADDMenuPop(PopUpMenu, '', '');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/04/2003
Modifié le ... :   /  /
Description .. : Renvoie le libellé par défaut d'un exercice
Mots clefs ... : EXERCICE;LIBELLE
*****************************************************************}

function LibelleExerciceDefaut(DateDebut, DateFin: TDateTime; bAbrege: boolean =
  False): string;
var
  AnneeDebut, AnneeFin: Word;
  Mois, Jour: Word;
begin
  DecodeDate(DateDebut, AnneeDebut, Mois, Jour);
  DecodeDate(DateFin, AnneeFin, Mois, Jour);
  if (AnneeFin = AnneeDebut) then
    Result := TraduireMemoire('Exercice') + ' ' + FormatDateTime('yyyy',
      DateDebut)
  else
  begin
    if bAbrege then
      Result := TraduireMemoire('Ex.') + ' ' + FormatDateTime('yyyy', DateDebut)
        + '/' + FormatDateTime('yyyy', DateFin)
    else
      Result := TraduireMemoire('Exercice') + ' ' + FormatDateTime('yyyy',
        DateDebut) + '/' + FormatDateTime('yyyy', DateFin);
  end;
end;

// FQ 12422

procedure TOuvExo.BFermeClick(Sender: TObject);
begin
  _DeblocageMonoPoste(True);
  Close;
  if IsInside(Self) then
    THPanel(parent).CloseInside;
end;

end.

