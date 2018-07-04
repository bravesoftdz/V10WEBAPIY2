{***********UNITE*************************************************
Auteur  ...... : FLO
Créé le ...... : 01/04/2008
Modifié le ... :   /  /
Description .. : Sélection des salariés pour convocation formation
Mots clefs ... : SALARIESCONVOC;CONVOCATION;FORMATION;
*****************************************************************
  N° |    Date    | Aut.| Version  |                Description
}
unit UTOFSalariesConvoc;

interface
uses
  Windows, Messages, SysUtils, Forms, HEnt1, HCtrls,
  {$IFNDEF EAGLCLIENT}  Db,  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} {$ENDIF}
  Classes, StdCtrls, UTOF;

type
  TOF_SALARIESCONVOC = class(TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
    Grille: THGrid;
    Mode  : String;
    Stage,Millesime,Session : String;
    procedure Valide     (Sender : TObject);
    procedure SelectAll  (Sender : TObject);
    Function  ChargeSalariesSession : Boolean;
    Procedure Refresh    (Sender : TObject);
  end;

implementation

uses HTB97,HMsgBox,HSysMenu,UTob,Vierge,PGOutilsFormation;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 01/04/2008
Modifié le ... :   /  /
Description .. : Chargement de l'écran
Mode gérés ... : SALARIES_SESSION
Mots clefs ... :
*****************************************************************}
procedure TOF_SALARIESCONVOC.OnArgument (S : String) ;
Begin
    Mode      := ReadTokenSt(S);
	If Mode = 'SALARIES_SESSION' Then
	Begin
		Stage     := ReadTokenSt(S);
		Millesime := ReadTokenSt(S);
		Session   := ReadTokenSt(S);
	End;

    Grille  := THGrid(GetControl('GRLISTESAL'));
    TToolbarButton97(GetControl('BValider')).OnClick   := Valide;
    TToolbarButton97(GetControl('BSelectAll')).OnClick := SelectAll;
    TCheckBox(GetControl('CKDEJARECU')).OnClick        := Refresh;

    Grille.ColAligns[2] := TaCenter;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 01/04/2008
Modifié le ... :   /  /
Description .. : Chargement de la liste des salariés
Mots clefs ... :
*****************************************************************}
procedure TOF_SALARIESCONVOC.OnLoad;
var Ret : Boolean;
Begin
    Ret := False;

	If Mode = 'SALARIES_SESSION' then
        Ret := ChargeSalariesSession;

	If Not Ret Then
    Begin
        TToolBarButton97(GetControl('BFerme')).Click;
        Exit;
    End;

    // Adaptation de la taille des colonnes
    TFVierge(Ecran).HMTrad.ResizeGridColumns (Grille);

	// Sélection de tous les salariés par défaut
	SelectAll(Nil);
    TToolbarButton97(GetControl('BSelectAll')).Down := True;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 01/04/2008
Modifié le ... :   /  /
Description .. : Retourne les éléments sélectionnés sur validation
Mots clefs ... :
*****************************************************************}
procedure TOF_SALARIESCONVOC.Valide(Sender: TObject);
var i             : integer;
    ListeElements : String;
begin
    // Retourne les éléments sélectionnés
    if (Grille.NbSelected=0) and (TToolbarButton97(GetControl('BSelectAll')).Down=False) then
    begin
        PGIBox(TraduireMemoire('Aucun élément sélectionné'),TFVierge(Ecran).Caption);
        Exit;
    end;

    // Pas de données dans la grille
    If (Grille.RowCount = 2) And (Grille.CellValues[0,1] = '') Then
    Begin
        Exit;
    End;
    
    // Parcours des éléments sélectionnés
    For i := 0 to Grille.RowCount-1 do
    begin
        If Grille.IsSelected(i) Then
        Begin
            ListeElements := ListeElements + '"' + Grille.CellValues[0,i] + '",';
        End;
    end;
    // Suppression de la dernière virgule
    If ListeElements <> '' Then ListeElements := Copy (ListeElements, 1, Length(ListeElements) - 1);

    // On retourne la liste des éléments
    If ListeElements <> '' Then TFVierge(Ecran).Retour := ListeElements;

    TToolBarButton97(GetControl('BFerme')).Click;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 01/04/2008
Modifié le ... :   /  /
Description .. : Bouton "Sélectionner tout"
Mots clefs ... :
*****************************************************************}
procedure TOF_SALARIESCONVOC.SelectAll(Sender: TObject);
begin
    Grille.AllSelected := Not Grille.AllSelected;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 01/04/2008
Modifié le ... :   /  /
Description .. : Charge la liste des salariés devant participer à une session de formation
Mots clefs ... :
*****************************************************************}
Function  TOF_SALARIESCONVOC.ChargeSalariesSession : Boolean;
Var TobData : TOB;
    Req     : String;
Begin
    Result := False;

    // Initialisation de la grille au départ
    Grille.RowCount := 2;
    Grille.Cells[0,1] := '';
    Grille.Cells[1,1] := '';
    Grille.Cells[2,1] := '';
    
	// Salariés inscrits à la session
	Req := 'SELECT PFO_SALARIE,(PFO_NOMSALARIE||" "||PFO_PRENOM) AS NOMSALARIE,MAX(YMA_STRDATE) AS DATECONVOC FROM FORMATIONS '+
            'LEFT JOIN YMAILS ON YMA_UTILISATEUR="'+USER_MAIL_FORMATION+'" AND YMA_FROM="'+Stage+Session+'" AND YMA_SUBJECT LIKE "SAL|%" AND YMA_DEST=PFO_SALARIE '+
            'WHERE PFO_CODESTAGE="'+Stage+'" AND PFO_ORDRE="'+Session+'" AND PFO_MILLESIME="'+Millesime+'" AND (PFO_ETATINSCFOR="" OR PFO_ETATINSCFOR="ATT" OR PFO_ETATINSCFOR="VAL")';

	// Salariés à qui on a déjà envoyé une convocation
	If GetControlText('CKDEJARECU') = 'X' Then
	Begin
		Req := Req + ' AND PFO_SALARIE NOT IN (SELECT YMA_DEST FROM YMAILS WHERE YMA_UTILISATEUR="'+USER_MAIL_FORMATION+'" AND YMA_FROM="'+Stage+Session+'" AND YMA_SUBJECT LIKE "SAL|%") ';
	End;
    Req := Req + 'GROUP BY PFO_SALARIE,PFO_NOMSALARIE,PFO_PRENOM';

	TobData := TOB.Create('$LesInscrits', Nil, -1);
	TobData.LoadDetailFromSQL(Req);

	// Affichage dans la grille
	If TobData.Detail.Count > 0 Then
    Begin
        TobData.PutGridDetail(Grille, False, False, 'PFO_SALARIE;NOMSALARIE;DATECONVOC');
        Result := True;
    End;

	FreeAndNil(TobData);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 01/04/2008
Modifié le ... :   /  /
Description .. : Refraîchit la liste
Mots clefs ... :
*****************************************************************}
Procedure TOF_SALARIESCONVOC.Refresh (Sender : TObject);
Begin
    If Mode = 'SALARIES_SESSION' Then ChargeSalariesSession;

    TToolbarButton97(GetControl('BSelectAll')).Down := False;
    Grille.AllSelected := False;
End;

initialization
  registerclasses([TOF_SALARIESCONVOC]);
end.
