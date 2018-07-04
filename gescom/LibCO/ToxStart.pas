unit ToxStart;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, ComCtrls, StdCtrls, ExtCtrls, HPanel, uTob, hmsgbox, Grids, Hctrls,
  DBTables, voirTob, ToxMoteur,
  UTox, AglToxRun, ToxFunctions,
  Spin, ParamSoc ;

type
  TGestionEchange = class(TForm)
    Panel1: TPanel;
    ButtonStopTox: TToolbarButton97;
    ButtonStartTox: TToolbarButton97;
    CompteRendu: TMemo;
    GroupBox1: TGroupBox;
    AfficheDonnee: TCheckBox;
    AfficheErreur: TCheckBox;
    Intervalle: TSpinEdit;
    Label1: TLabel;
    Confirmation: TCheckBox;
    procedure ButtonStartToxClick(Sender: TObject);
    procedure ButtonStopToxClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
  private
      { Déclarations privées }
    procedure ToxInformation(P: PToxEnveloppe ; var ATraiter: boolean);
    procedure ToxTraitement (P: PToxEnveloppe ; Index: integer ; TheTob: TOB; var Areecrire, Archiver, Rejeter: boolean);
    procedure ToxErreur     (CodeErreur: Integer; const MessageErreur: String);

  public
    { Déclarations publiques }
  end;

////////////////////////////////////////////////////////////////////////////////
// Procédures ou fonctions à exporter
////////////////////////////////////////////////////////////////////////////////
   procedure ToxGestionEchanges ;

implementation

{$R *.DFM}

procedure ToxGestionEchanges ( ) ;
var
  X : TGestionEchange;
begin
  X := TGestionEchange.Create ( Application ) ;
  X.ShowModal ;
  X.Free ;
end ;

////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Bouton de démarrage des échanges
////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TGestionEchange.ButtonStartToxClick(Sender: TObject);
var Delai           : integer ;
begin
    Delai           := Intervalle.Value ;

    Intervalle.Enabled     := False ;
    ButtonStopTox.enabled  := True;
    ButtonStartTox.Enabled := False;

    //AglStartTox ( 'TESTTOX', nil, ToxTraitement, nil, '', '', '', '', '', DemarreImmediat, Delai*60 );
end;

///////////////////////////////////////////////////////////////////////////////////////////////////////////
// Bouton d'arrêt des échanges
////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TGestionEchange.ButtonStopToxClick(Sender: TObject);
begin
  //if AglStatusTox then
  //begin
       AglStopTox ;
       ButtonStopTox.enabled  := False;
       ButtonStartTox.Enabled := True;
       Intervalle.Enabled     := True ;
  //end;
end;

///////////////////////////////////////////////////////////////////////////////
// Fonction appelée par AglStartTox
// Doit-on traitée la TOX ???
//          Valeur Retournée       True  : la TOX doit être traitée
//                                 False : la TOX sera traitée plus tard
///////////////////////////////////////////////////////////////////////////////
procedure TGestionEchange.ToxInformation ( P: PToxEnveloppe ; var ATraiter: boolean ) ;
begin
  CompteRendu.lines.add('   ') ;
  //CompteRendu.lines.add('   Arrivée d''un fichier TOX émis par ' + P^.Emetteur^.Libelle + ' le ' + DateToStr(P^.DateMsg)) ;
  CompteRendu.lines.add('   Arrivée d''un fichier TOX émis par ' + P^.Emetteur^.Libelle) ;
  CompteRendu.lines.add('   Site Destinataire : ' + P^.Destinataire^.Libelle) ;
  CompteRendu.lines.add('   ') ;

  ATraiter := True ;
  if Confirmation.checked then
  begin
    //if PGIAsk('Arrivée d''un fichier TOX émis par ' + P^.Emetteur^.Libelle + ' le ' + DateToStr(P^.DateMsg) , 'Confirmation')<>mrYes then
    if PGIAsk('Arrivée d''un fichier TOX émis par ' + P^.Emetteur^.Libelle  , 'Confirmation')<>mrYes then
    begin
      ATraiter := False ;
    end;
  end;

   CompteRendu.lines.add('   Sortie confirmation ') ;

end ;

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Fonction appelée par AglStartTox, permettant de traitée la TOX
// Argument
//         DateFichier               Date du fichier TOX
//         SiteEmetteur              site émetteur
//         TOX                       TOX à traiter
//         AReecrire                 Si on retourne True  -> La TOX est ré-ecrite et remise dans Corbeille Arrivée
//                                                  False -> Pas de réécriture et prise en compte de "AArchiver"
//         AArchiver                 Si on retourne True  -> TOX traitée, elle est copiée dans "Corbeille Traitée"
//                                                  False -> TOX rejetée, elle est copiée dans "Corbeille Rejet"
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TGestionEchange.ToxTraitement ( P: PToxEnveloppe ; Index: integer ; TheTob: TOB; var Areecrire, Archiver, Rejeter: boolean);
var EtatTox      : integer;
    NbTraitement : integer ;
begin

  CompteRendu.lines.add('  Traitement de la TOX émise par ' + P^.Emetteur^.Libelle + ' le ' + FormatDateTime('ddmmyyyy', P^.DateMsg)) ;
  CompteRendu.lines.add('   Site Destinataire : ' + P^.Destinataire^.Libelle) ;
  CompteRendu.lines.add('   ') ;

  EtatTox := TraitementDeLaTox (TheTob, CompteRendu, True, AfficheDonnee.checked, AfficheErreur.checked) ;

  if EtatTox = ToxARetraiter then
  begin
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
    if NbTraitement >= GetParamSoc ('SO_GCTOXNBTRT') then
    begin
      CompteRendu.lines.add('   ----> TOX traitée ' + IntToStr (NbTraitement) + ' fois. Recopie du fichier dans la corbeille des rejets')  ;
      EtatTox := ToxRejeter ;
    end;
  end;

  if EtatTox = ToxRejeter then
  begin
    AReecrire := False ;
    Archiver  := False ;
    Rejeter   := True  ;
    CompteRendu.lines.add('   ----> Fichier rejeté')  ;
  end
  else if EtatTox = ToxIntegrer then
  begin
    AReecrire := False ;
    Archiver  := True  ;
    Rejeter   := False ;
    CompteRendu.lines.add('   ----> Fichier intégré')  ;
  end else
  begin
    AReecrire := True ;
    Archiver  := False;
    Rejeter   := False;
    CompteRendu.lines.add('   ----> Fichier comportant des erreurs. Il sera retraité ultérieurement')  ;
  end;
end ;

procedure TGestionEchange.ToxErreur ( CodeErreur: Integer ; const MessageErreur: String ) ;
begin
  ShowMessage ( 'Message TOX(' + IntToStr ( CodeErreur ) + ') ' + MessageErreur ) ;
end ;

procedure TGestionEchange.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  //if AglStatusTox then AglStopTox ;
end;

procedure TGestionEchange.FormCreate(Sender: TObject);
begin
     AfficheDonnee.Checked := GetParamSoc ('SO_GCTOXAFFDONNEES');
     AfficheErreur.Checked := GetParamSoc ('SO_GCTOXAFFERREURS');
     Intervalle.Value      := GetParamSoc ('SO_GCTOXDELAI');
     Confirmation.Checked  := GetParamSoc ('SO_GCTOXCONFIRM');
end;

end.
