{***********UNITE*************************************************
Auteur  ...... : Franck VAUTRAIN
Créé le ...... : 28/06/2011
Modifié le ... :   /  /
Description .. : Class destiné à gérer les éditions dans une fiche de type
Suite ........ : autre que QR1
Mots clefs ... :
*****************************************************************}
unit UtilsEtat;

interface

uses  M3FP, StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOB,UTOF, AglInit, Agent,EntGC,
{$IFDEF EAGLCLIENT}
      MaineAGL,HPdfPrev,UtileAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
      EdtEtat,EdtDoc,
      EdtREtat,EdtRDoc,
{$ENDIF}
      HTB97;


Type
    TOptionEdition = class(Tobject)
    private
       fTip         : String;
       fNat         : String;
       fModele      : String;
       fTitre       : String;
       fstSQL       : String;
       fApercu      : Boolean;
       fDeuxPages   : Boolean;
       fFirst       : boolean;
       fListeExport : Boolean;
       fDuplicata   : Boolean;
       fCouleur     : Boolean;
       fSpages      : TPageControl;
       fEtat        : THValComboBox;
    public
      property    Tip         : String  read fTip         write fTip;
      property    Nature      : String  read fNat         write fNat;
      property    Modele      : String  read fModele      write fModele;
      property    Titre       : String  read fTitre       write fTitre;
      property    Apercu      : Boolean read fApercu      write fApercu;
      property    DeuxPages   : Boolean read fDeuxPages   write fDeuxPages;
      property    ListeExport : Boolean read fListeExport write fListeExport;
      property    Duplicata   : Boolean read fDuplicata   write fDuplicata;
      property    First       : Boolean read fFirst       write fFirst;
      property    Couleur     : Boolean read fFirst       write fCouleur;
      property    Spages      : TpageControl read fSpages write fSpages;
      //
      constructor Create (Tip, Nat, Modele, Titre, StSQL : String; Apercu, DeuxPages, First, ListeExport, Duplicata : Boolean; Spages : TPageControl; Etat : THValComboBox); overload;
      constructor Create;  overload;
      destructor  Destroy; override;
      //
      procedure   Appel_Generateur;
      procedure   ChargeListeEtat(VC: THValComboBox; var Idef : integer);
      function    LanceImpression(StSQL : string; TobEdition : TOB) : Integer;
    end;


implementation

constructor TOptionEdition.Create(Tip, Nat, Modele, Titre, StSQL : String; Apercu, DeuxPages, First, ListeExport, Duplicata : Boolean; Spages : TPageControl; Etat : THValComboBox);
begin

  fTip          := Tip;
  fNat          := Nat;
  fModele       := Modele;
  fTitre        := Titre;
  fstSQL        := StSQL;

  fApercu       := Apercu;
  fduplicata    := Duplicata;

  FListeExport  := ListeExport;
  fDeuxPages    := DeuxPages;
  fFirst        := First;
  fSpages       := Spages;

  fEtat         := Etat;
  fCouleur      := Couleur;

end;

constructor TOptionEdition.Create;
begin
  fTip          := '';
  fNat          := '';
  fModele       := '';
  fTitre        := '';
  fstSQL        := '';
  fApercu       := true;
  fduplicata    := false;
  FListeExport  := false;
  fDeuxPages    := false;
  fFirst        := true;
  fSpages       := nil;
  fEtat         := nil;
  //fCouleur      := false;
end;

procedure TOptionEdition.ChargeListeEtat( VC : THValComboBox; var Idef : integer) ;
Var QQ      : TQuery ;
    i_ind   : Integer ;
    StrSQL  : String ;
    OldDef  : String;
    Plus    : string;
BEGIN

    if fEtat = nil then Exit;

    if fFirst then iDef:=-1 else OldDef:=fEtat.Value;

    fEtat.Clear;

    Plus := fEtat.Plus;

    if plus <> '' then
      StrSQL:='SELECT MO_CODE, MO_LIBELLE, MO_DEFAUT FROM MODELES WHERE MO_TYPE="'+ fTip + '" ' + Plus      //'" AND MO_NATURE="'+ fNat +'"' ;
    else
      StrSQL:='SELECT MO_CODE, MO_LIBELLE, MO_DEFAUT FROM MODELES WHERE MO_TYPE="'+ fTip + '" AND MO_NATURE="'+ fNat +'"' ;

    QQ:=OpenSQL(StrSQL,TRUE) ;

    While Not QQ.EOF do
    BEGIN
      i_ind := fEtat.Items.Add(QQ.FindField('MO_LIBELLE').AsString) ;
      //si le modèle est chargé au moment de l'appel prendre le modèle chargé sinon prendre celui par
      //défaut en fonction de la nature...
      if fModele <> '' then
      begin
        if QQ.FindField('MO_CODE').AsString = fModele then Idef := i_ind;
      end
      else
      begin
        if QQ.FindField('MO_DEFAUT').AsString='X' then
        begin
         Idef   := i_ind ;
         fModele := QQ.findField('MO_CODE').AsString;
        end;
      end;
      fEtat.Values.Add(QQ.FindField('MO_CODE').AsString) ;
      QQ.Next
    END ;

    Ferme(QQ);

    if idef < 0 then Idef := 0;

    if fFirst then
      if idef>0 then fEtat.ItemIndex:=idef else fEtat.ItemIndex:=fEtat.Values.IndexOf(fmodele)
    else
      if fEtat.Values.IndexOf(OldDef)>=0 then fEtat.Value:=OldDef else fEtat.ItemIndex:=0 ;

END;

procedure TOptionEdition.Appel_Generateur;
var Idef : Integer;
BEGIN

  if not JaiLeDroitConcept(ccParamEtat, True) then Exit;

  if fTip='E' then
    EditEtat(fTip, fNat, fModele, true, fSpages, fstSQL, fTitre)
  else
    EditDocument(fTip,fNat,fModele,False);

  ChargeListeEtat(fEtat, Idef);

END;

function ToptionEdition.LanceImpression(StSQL : string; TobEdition : TOB) : Integer;
begin

  if Assigned(TobEdition) then
    Result := LanceEtatTob(fTip, fNat, fModele, TobEdition, fApercu, fListeExport, fDeuxPages, fSpages, fstSQL, fTitre, fDuplicata,1,'')
  else
    Result := LanceEtat(fTip, fNat, fModele, fApercu, fListeExport, fDeuxPages, fSpages, fstSQL, fTitre, fDuplicata);

end;

destructor TOptionEdition.Destroy;
begin
inherited;
end;



end.
