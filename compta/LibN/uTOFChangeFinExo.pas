unit uTOFChangeFinExo;

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
  Vierge,
  StdCtrls,
{$IFDEF EAGLCLIENT}

{$ELSE}
  DBTables,
  Hdb,
  DB,
{$ENDIF}
  Mask,
  uTob,         // TOB
  Hctrls,       // TableToPrefixe
  HSysMenu,
  HPanel,       // THPanel
  HTB97,        // TToolBarButton97
  HMsgBox,      // PgiError
  UiUtil,       // FindInsidePanel
  Ent1;         // TExoDate

procedure ChangementFinExercice;

type
  TFChangeFinExo = class(TFVierge)
    TLDATEFIN: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    TLLIBELLEEXO: THLabel;
    E_LIBELLE: THCritMaskEdit;
    HLabel1: THLabel;
    E_Exercice: THValComboBox;
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure E_ExerciceChange(Sender: TObject);
  private
    { Déclarations privées }
    FExerciceDepart  : string;    // Code Exercice de départ
    FExerciceArrivee : string;    // Code Exercice à mettre dans les écritures
    FNouvelleDate    : TDateTime; // Nouvelle Date de Fin
    FDateSuperieure  : TDateTime; //
    FDateInferieure  : TDateTime; //

    procedure MAJExercice;
    procedure MAJEcriture;
    procedure MAJAutreTable ( vStNomTable : string );
    function  ExisteEcriturePosterieure : Boolean;
    function  VerifIntegriteBordereau : Boolean;

  public
    { Déclarations publiques }
  end;

var FChangeFinExo: TFChangeFinExo;

implementation

uses uLibExercice,   // CRelatifVersExercice
     SoldeCpt,       // MajTotTousComptes
     HEnt1,          // BeginTrans
     uLibEcriture,   // CNumeroPiece
     uLibAnalytique; // CChargeAna

{$R *.DFM}

procedure ChangementFinExercice;
var FChangeFinExo : TFChangeFinExo;
    PP : THPanel;
begin
  PP := FindInsidePanel;
  FChangeFinExo := TFChangeFinExo.Create(Application);
  if PP = nil then
  begin
    try
      FChangeFinExo.ShowModal;
    finally
      FChangeFinExo.Free;
    end;
    Screen.Cursor := SyncrDefault;
  end
  else
  begin
    InitInside(FChangeFinExo, PP);
    FChangeFinExo.Show;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFChangeFinExo.BValiderClick(Sender: TObject);
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFChangeFinExo.BFermeClick(Sender: TObject);
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/06/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFChangeFinExo.E_ExerciceChange(Sender: TObject);
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/06/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFChangeFinExo.MAJExercice;
begin
  ExecuteSQL('UPDATE EXERCICE SET EX_DATEFIN = "' +
             UsDateTime( FNouvelleDate ) + '" WHERE ' +
             'EX_EXERCICE = "' + VH^.Encours.Code + '"');

  if VH^.Suivant.Code <> '' then
  begin
    ExecuteSQL('UPDATE EXERCICE SET EX_DATEDEBUT = "' +
               UsDateTime( FNouvelleDate + 1 ) + '" WHERE ' +
               'EX_EXERCICE = "' + VH^.Suivant.Code + '"');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TFChangeFinExo.ExisteEcriturePosterieure : Boolean;
begin
  Result := ExecuteSQL('SELECT E_JOURNAL FROM ECRITURE WHERE E_DATECOMPTABLE > "' +
            UsDateTime( FNouvelleDate ) + '" AND ' +
            'E_EXERCICE = "' + VH^.Encours.Code + '" ORDER BY E_JOURNAL') <> 0 ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFChangeFinExo.MAJEcriture;
var lQuery         : TQuery;
    lQueryMereEcr  : TQuery;
    lQueryMEreEcr2 : TQuery;
    lTobMereEcr    : Tob;
    lTobMEreEcr2   : Tob;
    lTobEcriture   : Tob;
    lTobAxe        : Tob;
    lTobAna        : Tob;
    lStSql         : string;
    lStSqlMereEcr  : string;
    lStSqlMereEcr2 : string;
    i,j,k          : integer;
    lIndice        : integer;

    function TrouvePiece( vPrefixe : string ; vCodeExercice : string ) : string;
    begin
      Result := vPrefixe + '_EXERCICE = "'+ vCodeExercice + '" AND ' +
                vPrefixe + '_JOURNAL = "' + lQuery.FindField('E_JOURNAL').AsString + '" AND ' +
                vPrefixe + '_PERIODE = ' + IntToStr(lQuery.FindField('E_PERIODE').AsInteger) + ' AND ' +
                vPrefixe + '_NUMEROPIECE = ' + IntToStr( lQuery.FindField('E_NUMEROPIECE').AsInteger) + ' AND ' +
                vPrefixe + '_QUALIFPIECE = "'+ lQuery.FindField('E_QUALIFPIECE').AsString + '" ';
    end;

begin
  // Récupération des Entêtes des Borderaux concernés
  lStSql := 'SELECT E_EXERCICE, E_JOURNAL, E_PERIODE, E_NUMEROPIECE, E_QUALIFPIECE FROM ECRITURE WHERE ' +
            '(E_MODESAISIE = "LIB" OR E_MODESAISIE = "BOR")' +
            ' AND E_EXERCICE = "' + FExerciceDepart + '"' +
            ' AND E_DATECOMPTABLE > "' + UsDateTime( FDateInferieure ) + '"' +
            ' AND E_DATECOMPTABLE <= "' + UsDateTime( FDateSuperieure ) + '"' +
            ' GROUP BY E_EXERCICE, E_JOURNAL, E_QUALIFPIECE, E_PERIODE, E_NUMEROPIECE';
  lQuery := nil;
  lTobMereEcr  := Tob.Create('', nil, -1);
  lTobMereEcr2 := Tob.Create('', nil, -1);
  try
    try
      lQuery := OpenSql( lStSql, True);
      while not lQuery.Eof do
      begin
        // Récupération complète du Bordereau
        lStSQLMereEcr := TrouvePiece( 'E', FExerciceDepart );

        lQueryMereEcr := OpenSql('SELECT * FROM ECRITURE WHERE ' + lStSQLMereEcr +
                         ' ORDER BY E_EXERCICE, E_JOURNAL, E_QUALIFPIECE, E_PERIODE, E_NUMEROPIECE, E_NUMLIGNE', True);

        // Requête qui récupère le bordereau d'arrivé des écritures à traiter
        lStSQLMereEcr2 := TrouvePiece( 'E', FExerciceArrivee );

        lQueryMereEcr2 := OpenSql( 'SELECT * FROM ECRITURE WHERE ' + lStSQLMereEcr2 +
                          ' ORDER BY E_EXERCICE, E_JOURNAL, E_QUALIFPIECE, E_PERIODE, E_NUMEROPIECE, E_NUMLIGNE', True);

        // Suppression des TobFilles
        lTobMereEcr.ClearDetail;
        lTobMereEcr2.ClearDetail;
        // Chargement du Bordereau
        lTobMereEcr.LoadDetailDB( 'ECRITURE', '', '', lQueryMereEcr,  False);
        lTobMereEcr2.LoadDetailDB('ECRITURE', '', '', lQueryMereEcr2, False);

        // Libération des requêtes
        Ferme( lQueryMereEcr  );
        Ferme( lQueryMereEcr2 );

        // Parcours du Bordereau pour changer l'exercice des écritures concernées
        i := lTobMereEcr.Detail.Count - 1;
        repeat
          lTobEcriture := lTobMereEcr.Detail[i];
          // Chargement de l'analytique de la pièce entière
          CChargeAna( lTobEcriture );

          if (lTobEcriture.GetValue('E_DATECOMPTABLE') <=  FDateSuperieure) and
             (lTobEcriture.GetValue('E_DATECOMPTABLE') > FDateInferieure) then
          begin
            // Indice pour l'ajout dans le ChangeParent
            lIndice := lTobMereEcr2.Detail.Count;

            for j := 0 to lTobEcriture.Detail.Count - 1 do
            begin
              lTobAxe := lTobEcriture.Detail[j];
              // Parcours des Sections Analytiques
              for k := 0 to lTobAxe.Detail.Count - 1 do
              begin
                // Parcours des Lignes d'Analytiques
                lTobAna := lTobAxe.Detail[k];
                lTobAna.PutValue('Y_EXERCICE', FExerciceArrivee);
              end;
            end;

            // Ecritures qui changent d'exercice
            lTobEcriture.PutValue('E_EXERCICE', FExerciceArrivee);
            lTobEcriture.PutValue('E_IO', 'X');
            lTobEcriture.ChangeParent( lTobMereEcr2, lIndice );
          end;
          i := i - 1;
        until (i < 0);

        // Re-numérotation de la pièce
        CNumeroPiece( lTobMereEcr  );
        //lTobMereEcr.SaveToFile('C:\LTOBMEREECR.TXT', False, True, True);
        CNumeroPiece( lTobMereEcr2 );
        //lTobMereEcr2.SaveToFile('C:\LTOBMEREECR2.TXT', False, True, True);

        // Suppression en Base des Bordereaux avant ré-insertion
        ExecuteSQL('DELETE FROM ECRITURE WHERE ' + lStSQLMereEcr);
        ExecuteSQL('DELETE FROM ECRITURE WHERE ' + lStSQLMereEcr2);

        // Suppression en Base de l'analytique
        ExecuteSQL('DELETE FROM ANALYTIQ WHERE ' + TrouvePiece( 'Y', FExerciceDepart));
        ExecuteSQL('DELETE FROM ANALYTIQ WHERE ' + TrouvePiece( 'Y', FExerciceArrivee));

        // Enregistrement en Base
        lTobMereEcr.InsertDb ( nil, False );
        lTobMereEcr2.InsertDb( nil, False );

        // Passage à l'entête de Bordereau suivante
        lQuery.Next;
      end;
    except
      on E: Exception do
      begin
        PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : MAJEcriture');
        Raise;
      end;
    end;

  finally
    Ferme( lQuery );
    // Libération Mémoire des Tob
    FreeAndNil( lTobMereEcr  );
    FreeAndNil( lTobMereEcr2 );
  end;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/06/2004
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFChangeFinExo.MAJAutreTable(vStNomTable: string);
var lStPrefixe : string;
begin
  if vStNomTable = '' then Exit;
  lStPrefixe := TableToPrefixe( vStNomTable );
  ExecuteSQL('UPDATE ' + vstNomTable + ' SET ' + lStPrefixe + '_EXERCICE = "' + FExerciceArrivee + '" WHERE ' +
             lStPrefixe + '_EXERCICE = "' + VH^.Encours.Code + '" AND ' +
             lStPrefixe + '_DATECOMPTABLE <= "' + UsDateTime( FNouvelleDate ) + '"');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/06/2004
Modifié le ... :   /  /    
Description .. : Vérifie que l'équilibre Folio d'un bordereau de type libre
Suite ........ : n'est pas perdu lors du changement du changement de la Date
Mots clefs ... :
*****************************************************************}
function TFChangeFinExo.VerifIntegriteBordereau : Boolean;
var lSt : string;
    lQuery : TQuery;
begin
  if FNouvelleDate > VH^.Encours.Fin then
    lSt := 'SELECT SUM(E_DEBIT)-SUM(E_CREDIT) TOTAL FROM ECRITURE WHERE E_MODESAISIE = "LIB" AND ' +
           'E_EXERCICE = "' + VH^.Encours.Code + '" AND ' +
           'E_DATECOMPTABLE > "' + UsDateTime( VH^.Encours.Fin ) + '" AND ' +
           'E_DATECOMPTABLE <= "' + UsDateTime( FNouvelleDate ) + '" ' +
           'GROUP BY E_EXERCICE, E_JOURNAL, E_QUALIFPIECE, E_PERIODE, E_NUMEROPIECE'
  else
    lSt := 'SELECT SUM(E_DEBIT)-SUM(E_CREDIT) TOTAL FROM ECRITURE WHERE E_MODESAISIE = "LIB" AND ' +
           'E_EXERCICE = "' + VH^.Encours.Code + '" AND ' +
           'E_DATECOMPTABLE > "' + UsDateTime( FNouvelleDate ) + '" AND ' +
           'E_DATECOMPTABLE <= "' + UsDateTime( VH^.Encours.Fin ) + '" ' +
           'GROUP BY E_EXERCICE, E_JOURNAL, E_QUALIFPIECE, E_PERIODE, E_NUMEROPIECE';

  try
    lQuery := OpenSQL( lSt, True);
    Result := lQuery.FindField('TOTAL').AsFloat = 0;
  finally
    Ferme( lQuery );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
end.



