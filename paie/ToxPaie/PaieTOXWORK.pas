unit PaieTOXWORK;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
  uTob,
  PaieToxMoteur,
  ParamSoc,
  HTB97,
  ExtCtrls,
  HPanel,
  uTox,
  Hent1,
  uToxConst,
  uToxClasses,
  uToxDecProc ;

procedure PaieNotConfirmeTox ( P: TCollectionEnveloppe; var Atraiter : boolean) ;
procedure PaieTraitementTox  ( P: TCollectionEnveloppe ; Index: integer ; TheTob: TOB ; var ttResult: TToxTraitement ; ToFeel: TToxFeelTob = Nil ) ;
procedure ToxTraitementSansEcran   ( P: TCollectionEnveloppe ; TheTob: TOB ; var Resultat: TToxTraitement);

type
  TTFicheTOXWork = class(TForm)
    CompteRendu: TMemo;
    procedure FormActivate(Sender: TObject);
  private
    { Déclarations privées }
    procedure ToxTraitement ;
  public
    { Déclarations publiques }
    P: TCollectionEnveloppe ;
    Index: integer ;
    TheTob: TOB;
    Resultat: TToxTraitement ;
  end;


implementation

{$R *.DFM}

////////////////////////////////////////////////////////////////////////////////
// Fonction permettant de ne pas demander une confirmation de traitement
// a l'arrivée d'un fichier TOX
////////////////////////////////////////////////////////////////////////////////
procedure PaieNotConfirmeTox ( P: TCollectionEnveloppe; var Atraiter : boolean) ;
begin
  Atraiter := True ;
end;

////////////////////////////////////////////////////////////////////////////////
// Chainâge sur l'intégration d'un fichier TOX
////////////////////////////////////////////////////////////////////////////////
procedure PaieTraitementTox ( P: TCollectionEnveloppe; Index: integer ; TheTob: TOB ; var ttResult: TToxTraitement ; ToFeel: TToxFeelTob = Nil ) ;
var
  X: TTFicheTOXWork;
begin
  //
  // Cas 1 : Affichage Ecran de l'intégration
  //
{  if GetParamSoc ('SO_GCTOXCONFIRM') = True then
  begin
    X := TTFicheTOXWork.Create ( Application ) ;
    X.P      := P ;
    X.Index  := Index ;
    X.TheTob := TheTob ;
    X.ShowModal ;
    ttResult  := X.Resultat ;
    X.Free ;
  end else
  begin}
    ToxTraitementSansEcran ( P, TheTob, ttResult);
//  end;
end ;

procedure TTFicheTOXWork.ToxTraitement ;
var EtatTox      : integer ;
    NbTraitement : integer ;
    AfficheDonnee: boolean ;
    AfficheErreur: boolean ;
    Devise       : string  ;
    BasculeEuro  : boolean ;
begin
{  BasculeEuro:=True ;
  //
  // Récupération de la devise d'émission
  //
  Devise:=P.GetInfoComp('DEVISE', Devise);
  if Devise = '' then  Devise := V_PGI.DevisePivot ;
  //
  // Le site émetteut a t il basculé en EURO ?
  //
  BasculeEuro:=P.GetInfoComp('BASCULEEURO', BasculeEuro);
  //
  // Affichage des données intégrées ??? et des erreurs ???
  //
  AfficheDonnee := GetParamSoc ('SO_GCTOXAFFDONNEES');
  AfficheErreur := GetParamSoc ('SO_GCTOXAFFERREURS');
  //
  // Affichage message infos pour début d'intégration
  //
  CompteRendu.lines.add('  Traitement de la TOX émise par ' + P.eLibelle + ' le ' + FormatDateTime('ddmmyyyy', P.DateMsg)) ;
  CompteRendu.lines.add('  Site Destinataire : ' + P.dLibelle) ;
  CompteRendu.lines.add('   ') ;
  //
  //  Intégration de la TOX
  //
  EtatTox := TraitementDeLaTox (TheTob, CompteRendu, True, AfficheDonnee, AfficheErreur, False, True, Devise, BasculeEuro) ;
  //
  // MAJ du nombre de traitements
  //
  if not TheTob.FieldExists ('NB_TRAITEMENT') then
  begin
    TheTOB.AddChampSup ('NB_TRAITEMENT', False) ;
    TheTOB.PutValue    ('NB_TRAITEMENT', 1);
    NbTraitement := 1 ;
  end else
  begin
    NbTraitement := TheTOB.GetValue ('NB_TRAITEMENT');
    inc (NbTraitement) ;
    TheTOB.PutValue ('NB_TRAITEMENT', NbTraitement);
  end;
  //
  // Que faut -il faire de la TOX ????
  //
  if EtatTox = ToxARetraiter then
  begin
    //
    // Si Nbre de traitement > Nombre maximum de traitement, la TOX est définitivement rejetée
    //
    if NbTraitement >= GetParamSoc ('SO_GCTOXNBTRT') then
    begin
      CompteRendu.lines.add('   ----> TOX traitée ' + IntToStr (NbTraitement) + ' fois. Recopie du fichier dans la corbeille des rejets')  ;
      Resultat := [ ttRewrite, ttBad ] ;
    end else
    begin
      CompteRendu.lines.add('   ----> Fichier comportant des erreurs. Il sera retraité ultérieurement')  ;
      Resultat := [ ttRewrite, ttRetry ] ;
    end;
  end
  else if EtatTox = ToxRejeter then
  begin
    CompteRendu.lines.add('   ----> Fichier rejeté')  ;
    //Resultat := [ ttRewrite, ttBad ] ;
    Resultat := [ ttBad ] ;
  end
  else if EtatTox = ToxIntegrer then
  begin
    CompteRendu.lines.add('   ----> Fichier intégré')  ;
    Resultat := [ ttRewrite, ttOK ] ;
  end;
}
end ;

procedure ToxTraitementSansEcran ( P: TCollectionEnveloppe; TheTob: TOB ; var Resultat: TToxTraitement);
var EtatTox      : integer ;
    NbTraitement : integer ;
    Devise       : string  ;
    BasculeEuro  : boolean ;
begin
  NbTraitement := 0    ;
  //
  // Traitement de la TOX
  //
  EtatTox := PaieTraitementDeLaTox (TheTob, Tmemo(nil), False, False, False, False) ;
  //
  // MAJ du nombre de traitements
  //
  if not TheTob.FieldExists ('NB_TRAITEMENT') then
  begin
    TheTOB.AddChampSup ('NB_TRAITEMENT', False) ;
    TheTOB.PutValue    ('NB_TRAITEMENT', 1);
  end else
  begin
    NbTraitement := TheTOB.GetValue ('NB_TRAITEMENT');
    inc (NbTraitement) ; TheTOB.PutValue ('NB_TRAITEMENT', NbTraitement);
  end;
  //
  // Que fait on de la TOX ????
  //
  if EtatTox = ToxARetraiter then
  begin
    //
    // Si Nbre de traitement > Nombre maximum de traitement, la TOX est définitivement rejetée
    //
    if NbTraitement >= 50 then  Resultat := [ ttRewrite, ttBad ]
//    GetParamSoc ('SO_GCTOXNBTRT') then  Resultat := [ ttRewrite, ttBad ]
    else Resultat := [ ttRewrite, ttRetry ] ;
  end
  else if EtatTox = ToxRejeter  then Resultat := [ ttBad ]
  else if EtatTox = ToxIntegrer then Resultat := [ ttRewrite, ttOK ] ;
end ;

procedure TTFicheTOXWork.FormActivate(Sender: TObject);
begin
   ToxTraitement ;
end;

end.
