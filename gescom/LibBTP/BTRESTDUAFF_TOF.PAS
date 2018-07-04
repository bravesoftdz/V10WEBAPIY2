{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/02/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRESTDUAFF ()
Mots clefs ... : TOF;BTRESTDUAFF
*****************************************************************}
unit BTRESTDUAFF_TOF;

interface

uses StdCtrls,
     Controls,
     Classes,
     HTB97,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ELSE}
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     Menus,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     Stat,
     HQry,
     UTOB,
     utobview,
     AGLInit,
     facture,
     AffaireUtil,
     SaisUtil,
     FactCpta,
     EntGc,
     Ent1,
     Paramsoc,
     UtilReglementAffaire,
     uTofAfBaseCodeAffaire;

type
  TOF_BTRESTDUAFF = class(TOF_AFBASECODEAFFAIRE)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;

  private
    TV1           : TTobViewer;

    TOBRDA        : TOB;

    DateDeb       : TdateTime;
    DateFin       : TDateTime;

    CodeAffaire   : String;
    CodeAffaire1  : String;
    CodeAffaire2  : String;
    CodeAffaire3  : String;
    Tiers         : String;
    TypeAffaire   : string;
    NaturePiece   : String;

    THVAffaire0   : THValComboBox;
    Domaine       : THValComboBox;
    THVNatPiece   : THMultiValComboBox;

    THAffaire     : THEdit;
    THAffaire0    : THEdit;
    THAffaire1    : THEdit;
    THAffaire2    : THEdit;
    THAffaire3    : THEdit;
    THAvenant     : THEdit;

    THTiers       : THedit;

    TLAffaire     : THLabel;
    TLTiers       : THLabel;

    BSelectAff1 	: TToolBarButton97;
    BEffaceAff1 	: TToolBarButton97;

    procedure OnDblClickTV(Sender: TObject);
    procedure ChargeLesFactures;
    procedure CompletelesInfos;

    //Procedure de gestion des zones
    procedure AffaireChange(Sender: TObject);
    procedure TiersExit(Sender: TObject);

  end;

implementation

procedure TOF_BTRESTDUAFF.OnNew;
begin
  inherited;
end;

procedure TOF_BTRESTDUAFF.OnDelete;
begin
  inherited;
end;

procedure TOF_BTRESTDUAFF.OnUpdate;
begin
  inherited;

  ChargeLesFactures;

  CompletelesInfos;

  if TOBRDA <> nil then TFStat(Ecran).LaTOB := TOBRDA;

end;

procedure TOF_BTRESTDUAFF.OnLoad;
begin
  inherited;

  DateDeb       := StrtoDate(GetControltext('DATEPIECE'));
  DateFin       := StrtoDate(GetControltext('DATEPIECE_'));

  CodeAffaire   := THAffaire.text;
  CodeAffaire1  := THAffaire1.text;
  CodeAffaire2  := THAffaire2.text;
  CodeAffaire3  := THAffaire3.text;

  Tiers         := ThTiers.Text;

  TypeAffaire   := THVAffaire0.Value;
  THAffaire0.text := TypeAffaire;

  NaturePiece   := THVNatPiece.Text;

end;

procedure TOF_BTRESTDUAFF.OnArgument(S: string);
Var LastDay : Integer;
    WYear   : Word;
    WMonth  : Word;
    WDay    : Word ;

begin
  inherited;

  // Gestion des évènements du TB
  TV1 := TTobViewer(GetControl('TV'));
  TV1.OnDblClick := OnDblClickTV;

  //Chargement date debut mois et date fin mois en fonction date du jour
  DecodeDate(Now , WYear , WMonth , WDay);
  LastDay := DaysPerMonth(WYear, WMonth);

  DateDeb := EncodeDate(WYear, WMonth, 1);
  DateFin := EncodeDate(WYear, WMonth, LastDay);

  SetControlText('DATEPIECE', DateToStr(DateDeb));
  SetControlText('DATEPIECE_', DateToStr(DateFin));

  //chargement des zones de l'affaire
  THVAffaire0:=THValComboBox(GetControl('AFFAIRE0'));
  THVNatPiece:=THMultiValComboBox(GetControl('GP_NATUREPIECEG'));
//  THVNatpiece.Plus := ' AND GP_NATUREPIECEG IN ("FBT","FAC","ABT","AVC")';

  THAffaire  := THEdit(GetControl('AFF_AFFAIRE'));

  THAffaire0 := THEdit(GetControl('AFF_AFFAIRE0'));
  THAffaire1 := THEdit(GetControl('GP_AFFAIRE1'));
  THAffaire2 := THEdit(GetControl('GP_AFFAIRE2'));
  THAffaire3 := THEdit(GetControl('GP_AFFAIRE3'));
  THAvenant  := THEdit(GetControl('GP_AVENANT'));

  THVAffaire0.OnChange := AffaireChange;

  TLAffaire  :=THLabel(GetControl('TGP_AFFAIRE'));

  THTiers    := THEdit(GetControl('GP_TIERS'));
  THTiers.OnExit := TiersExit;

  TLTiers     := THLabel(GetControl('CT_TIERS'));

  BSelectAff1 := TToolBarButton97 (GetControl('BSELECTAFF1'));
  BEffaceAff1 := TToolBarButton97 (GetControl('BEFFACEAFF1'));

  //Initialisation des Zones écran
  THAffaire0.Text := 'A';
  THVAffaire0.Value := THAffaire0.Text;

  //Création Tob
  TOBRDA := TOB.create('Ma Tob', nil, -1);

  AffaireChange(self);

  if GetControl('GP_DOMAINE') IS THValComboBox then Domaine:=THValComboBox(GetControl('GP_DOMAINE'));
  if Domaine<>Nil then PositionneDomaineUser(Domaine) ;

end;

procedure TOF_BTRESTDUAFF.OnClose;
begin
  inherited;
  if TOBRDA <> nil then FreeAndNil(TOBRDA);
end;

//******************************************************************************
//********************* Zooms / double clic  ***********************************
//******************************************************************************

procedure TOF_BTRESTDUAFF.OnDblClickTV(Sender: TObject);
begin
end;

procedure TOF_BTRESTDUAFF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit);
begin

  Aff := THEdit(GetControl('AFF_AFFAIRE'));

  Aff0 := THEdit(GetControl('AFF_AFFAIRE0'));

  Aff1 := THEdit(GetControl('GP_AFFAIRE1'));
  Aff2 := THEdit(GetControl('GP_AFFAIRE2'));
  Aff3 := THEdit(GetControl('GP_AFFAIRE3'));
  Aff4 := THEdit(GetControl('GP_AVENANT'));
  Tiers := THEdit(GetControl('GP_TIERS'));
  
end;

procedure TOF_BTRESTDUAFF.ChargeLesFactures;
var
  Req : string;
  QQ  : Tquery;
begin

  Req := TFStat(Ecran).stSQL;

  //On met à jour la requete de la tob associée
  //if (NaturePiece = '') Or (NaturePiece = '<<Tous>>') then
  //   Req := Req + ' AND GP_NATUREPIECEG IN ("FBT","FAC","ABT","AVC") ';

  // On met à jour la TOB uniquement pour les chantiers
  // ayant des mouvements dans les dates sélectionnées
  if (DateToStr(DateDeb) <> '') and (DateToStr(DateFin) <> '') then
     Begin
     Req := Req + ' AND GP_DATEPIECE BETWEEN "' + UsDateTime(DateDeb);
     Req := Req + '" AND "' + UsDateTime(DateFin) + '"';
     end;

  QQ := OpenSQL(Req, true);

  TOBRDA.LoadDetailDB('', '', '', QQ, false);

  ferme(QQ);

end;

procedure TOF_BTRESTDUAFF.CompletelesInfos;
var I           : integer;
    Ecr         : RMvt;
    MtReglement : Double;
    MtFacture   : Double;
    MtAcompte   : Double;
    PourcentAvc : Double;
    DateEch     : TdateTime;
    DateRegl    : TDateTime;
    Auxiliaire  : String;
    Sql         : String;
    QQ          : TQuery;
begin

  if TOBRDA = nil then exit;
  if TOBRDA.detail.count = 0 then exit;
  i := 0;
  Repeat
      Begin
      //Chargement de la dernière date d'échéance sur facture
      DateEch := Idate1900;
      SQL := 'SELECT GPE_DATEECHE FROM PIEDECHE WHERE GPE_NATUREPIECEG = "'  + TOBRDA.Detail[i].GetValue('GP_NATUREPIECEG') + '" ';
      SQL := SQL + 'AND GPE_NUMERO=' +  IntToStr(TOBRDA.Detail[i].GetValue('GP_NUMERO'));
      QQ := OpensQL(SQL, True);
      if not QQ.Eof then
         Begin
         DateEch := QQ.findfield('GPE_DATEECHE').AsDateTime;
         end;
			if (TCheckBox(GetCOntrol('FACTURESECHUES')).Checked) and (DateEch >= V_PGI.DateEntree) then
      	Begin
        TOBRDA.Detail[i].Free; // On ne prend pas en compte les factures non arrivées à échéance
      	End Else
        Begin
        TOBRDA.Detail[i].AddChampSupValeur('DATEECHF', DateEch);
        //Création des champs supplémentaire par ligne de facture
        MtReglement := 0;
        MtAcompte   := 0;
        PourcentAvc := 0;
        MtFacture   := TOBRDA.Detail[i].GetValue('GP_TOTALTTCDEV');
        //Recherche du montant Réglement;, Montant acompte, Date Echéance et Date Réglement
        if TOBRDA.Detail[i].GetValue('GP_REFCOMPTABLE') <> '' then
           begin
           Ecr:=DecodeRefGCComptable(TOBRDA.Detail[i].GetValue('GP_REFCOMPTABLE')) ;
           Auxiliaire := TOBRDA.Detail[i].GetValue('T_AUXILIAIRE');
           //FV1 - 13/11/2014 : FS#1302 - DURET : Rien dans "Reste dû par Affaire"
           if GetParamSocSecur('SO_BTCOMPTAREGL', False) then
              GetInfosReglementsPiece(ECR, Auxiliaire, MtReglement, MtAcompte, DateEch, DateRegl)
           else
              GetInfosReglementsAcompte(TOBRDA.Detail[i], MtReglement, MtAcompte, DateEch, DateRegl);
           end;
        if (MtReglement > MtFacture) then MtReglement := MtFacture; //dans le cas où le reglement aurait été enregistré pour plusieurs factures
        TOBRDA.Detail[i].AddChampSupValeur('MONTANTREG', MtReglement);
        TOBRDA.Detail[i].AddChampSupValeur('MONTANTACT', MtAcompte);
        TOBRDA.Detail[i].AddChampSupValeur('DATEREG', DateRegl);
        TOBRDA.Detail[i].AddChampSupValeur('RESTEDU', MtFacture-MtReglement-MtAcompte);
        TOBRDA.Detail[i].AddChampSupValeur('DATEECHR', DateEch);
        //Calcul du Pourcentage d'avancement en fonction des lignes document
        SQL := 'SELECT SUM(GL_QTESIT*GL_PUHTDEV) as QTEPUHT, SUM(GL_QTEPREVAVANC*GL_PUHTDEV) AS QTEPREV ';
        SQL := SQL + 'FROM LIGNE WHERE GL_NATUREPIECEG = "' + TOBRDA.Detail[I].GetValue('GP_NATUREPIECEG') + '" ';
        SQL := SQL + 'AND GL_NUMERO=' +  IntToStr(TOBRDA.Detail[i].GetValue('GP_NUMERO'));
        SQL := SQL + ' AND GL_TYPELIGNE="ART"';
        QQ := OpensQL(SQL, True);
        if not QQ.eof then
          Begin
          if QQ.findfield('QTEPREV').AsFloat <> 0 then
            PourcentAvc := (QQ.findfield('QTEPUHT').AsFloat/QQ.findfield('QTEPREV').AsFloat) * 100;
          end;
        Ferme(QQ);
        TOBRDA.Detail[i].AddChampSupValeur('POURAVANC', PourcentAvc);
        Inc(i);
      	End;
      End;
  Until i >= TOBRDA.Detail.Count;
end;

Procedure TOF_BTRESTDUAFF.AffaireChange(Sender: TObject);
Begin

  //Gestion des zones
  if THVAffaire0.Value  = 'I' then
     Begin
     TLAffaire.Caption := 'Contrat : ';
     THAffaire0.text := 'A';
     SETCONTROLTEXT('GP_NATUREPIECEG', 'FAC');
     End
  Else if THVAffaire0.Value = 'W' then
     Begin
     TLAffaire.Caption := 'Appel : ';
     THAffaire0.text := 'W';
     SETCONTROLTEXT('GP_NATUREPIECEG', 'FAC');
     End
  Else
     Begin
     TLAffaire.Caption := 'Chantier : ';
     THAffaire0.text := 'A';
     SETCONTROLTEXT('GP_NATUREPIECEG', 'FBT');
     end;
  //
  THAffaire.Text := THVAffaire0.Value;
  //
  ChargeCleAffaire(THAffaire0, THAffaire1, THAffaire2, THAffaire3, THAvenant, BSelectAFF1, TaModif, THAffaire.Text, false);

end;

Procedure TOF_BTRESTDUAFF.TiersExit(Sender: TObject);
begin

  //Recherche du nom du tiers dans la table.
  TLTiers.caption := RechDom ('GCTIERS',THTiers.text,false);

end;

initialization
  registerclasses([TOF_BTRESTDUAFF]);
end.
