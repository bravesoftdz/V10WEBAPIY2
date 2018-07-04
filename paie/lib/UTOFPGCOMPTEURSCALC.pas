{***********UNITE*************************************************
Auteur  ...... : FLO
Créé le ...... : 26/06/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : COMPTEURSCALC_MUL ()
Mots clefs ... : TOF;PGCOMPTEURSCALC_MUL
*****************************************************************
PT1  21/08/2007  FLO  Ajout d'une sélection sur les compteurs calculés/à recalculer
PT2  08/02/2008  NA   Suppression en masse des compteurs calculés
}
unit UTOFPGCOMPTEURSCALC;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  DBGrids, HDB, Mul, Fiche, db,
{$IFNDEF DBXPRESS}
  dbTables,
{$ELSE}
  uDbxDataSet,
{$ENDIF}
  FE_Main,
{$ENDIF}
{$IFDEF EAGLCLIENT}
  eMul,MaineAGL,
{$ENDIF}
  sysutils, ComCtrls, HCtrls, HEnt1, UTOF, Hqry, HTB97, HMsgBox, P5Def;

type
  TOF_PGCOMPTEURSCALC_MUL = class(TOF)
    procedure OnArgument(S: string); override;
    procedure OnLoad               ; override ;  // PT1
  private
    procedure GrilleDblClick(Sender: TObject);
    procedure CreerCompteur(Sender: TObject);
    procedure SupprimeCompteurcalc(Sender: TObject); //PT2
  end;

  TOF_PGANALYSECOMPTEURS = class(TOF)
    procedure OnArgument(S: string); override;
  end;

implementation

procedure TOF_PGCOMPTEURSCALC_MUL.OnArgument(S: string);
var
  Liste: THGrid;
  BNouv: TToolBarButton97;
  Num: Integer;
begin
  inherited;

     // Gestion du double-clic dans la liste
  Liste := THGrid(GetControl('FLISTE'));
  if Liste <> nil then Liste.OnDblClick := GrilleDblClick;

     // Gestion du bouton nouveau
  BNouv := TToolbarButton97(GetControl('BInsert'));
  if BNouv <> nil then BNouv.OnClick := CreerCompteur;

     // pt3 Suppression en masse des compteurs calculés
  (GetControl('B_SUPPRIMER') as TToolBarButton97).Onclick := SupprimeCompteurcalc; //PT2

     // Génération des dates de début et de fin correspondant au mois en cours
  SetControlText('PYP_DATEDEBUTPRES', DateToStr(DebutdeMois(Date())));
  SetControlText('PYP_DATEFINPRES', DateToStr(FindeMois(Date())));

     // Recherche des libellés des zones TRAVAIL, CODESTAT, LIBREPCMB
  for Num := 1 to 4 do
  begin
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PYP_TRAVAILN' + IntToStr(Num)), GetControl('TPYP_TRAVAILN' + IntToStr(Num)));
  end;

  VisibiliteStat(GetControl('PYP_CODESTAT'), GetControl('TPYP_CODESTAT'));

  for Num := 1 to 4 do
  begin
    VisibiliteChamplibresal(Inttostr(Num), GetControl('PYP_LIBREPCMB' + IntToStr(Num)), GetControl('TPYP_LIBREPCMB' + IntToStr(Num)));
  end;

  // Compteurs Calculés par défaut
  SetControlText('CBAFFICHAGE', 'CAL'); //PT1
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 21/08/2007 / PT1
Modifié le ... :   /  /    
Description .. : Chargement de la liste
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGCOMPTEURSCALC_MUL.OnLoad ;
begin
     If GetControlText('CBAFFICHAGE') = 'CAL' Then
     Begin
          SetControlText('XX_WHERE', ' AND PYP_PGINDICATPRES<>"ARE"');
          TFmul(Ecran).Caption := 'Compteurs de présence calculés';
          UpdateCaption(TFmul(Ecran));
     End
     Else
     Begin
          SetControlText('XX_WHERE', ' AND PYP_PGINDICATPRES="ARE"');
          TFmul(Ecran).Caption := 'Compteurs de présence à recalculer';
          UpdateCaption(TFmul(Ecran));
     End;
end ;

procedure TOF_PGCOMPTEURSCALC_MUL.GrilleDblClick(Sender: TObject);
var Q_Mul: THQuery;
  St: string;
{$IFDEF EAGLCLIENT}
  Liste: THGrid;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  Liste := THGrid(GetControl('FLISTE'));
  TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
{$ENDIF}
  Q_Mul := THQuery(Ecran.FindComponent('Q'));

  St := Q_Mul.FindField('PYP_SALARIE').AsString + ';' + Q_Mul.FindField('PYP_DATEDEBUTPRES').AsString + ';' + Q_Mul.FindField('PYP_DATEFINPRES').AsString + ';' + Q_Mul.FindField('PYP_COMPTEURPRES').AsString + ';' + Q_Mul.FindField('PYP_TYPECALPRES').AsString;

  AGLLanceFiche('PAY', 'COMPTEURSCALC', St, St, 'ACTION=MODIFICATION;' + Q_Mul.FindField('PYP_COMPTEURPRES').AsString + ';' + Q_Mul.FindField('DATEVALIDITE').AsString + ';' + Q_Mul.FindField('PYR_LIBELLE').AsString);

     // Rafraîchissement de la liste
  if GetControl('BCherche') is TToolbarButton97 then (GetControl('BCherche') as TToolbarButton97).Click;
end;

procedure TOF_PGCOMPTEURSCALC_MUL.CreerCompteur(Sender: TObject);
begin
  AGLLanceFiche('PAY', 'COMPTEURSCALC', '', '', 'ACTION=CREATION');
     // Rafraîchissement de la liste
  if GetControl('BCherche') is TToolbarButton97 then (GetControl('BCherche') as TToolbarButton97).Click;
end;

// pt2
{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 08/02/2008
Modifié le ... :   /  /
Description .. : Suppression en masse des compteurs calculés
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGCOMPTEURSCALC_MUL.Supprimecompteurcalc(Sender: TObject);
Var Salarie, compteur, typecalpres : String;
    i : Integer;
    DateDebpres,DateFinpres : TDateTime;
begin
     // Elements sélectionnés
     If (TFmul(Ecran).Fliste.NbSelected = 0) And (Not TFmul(Ecran).Fliste.AllSelected) Then
     Begin
          PGIBox(TraduireMemoire('Aucune ligne n''est sélectionnée.'), TraduireMemoire('Suppression des compteurs calculés'));
          Exit;
     End;

     If PgiAsk(TraduireMemoire('Voulez-vous supprimer les compteurs calculés sélectionnés ?'), Ecran.caption) = mrYes Then
     Begin


          // Si tout est sélectionné
          If (TFmul(Ecran).Fliste.AllSelected = True) Then
          Begin
               TFmul(Ecran).Q.First;
               While Not TFmul(Ecran).Q.EOF Do
               Begin
                 Salarie := TFmul(Ecran).Q.FindField('PYP_SALARIE').AsString;
                 datedebpres := TFmul(Ecran).Q.FindField('PYP_DATEDEBUTPRES').AsDatetime;
                 datefinpres := TFmul(Ecran).Q.FindField('PYP_DATEFINPRES').AsDatetime;
                 compteur := TFmul(Ecran).Q.FindField('PYP_COMPTEURPRES').AsString;
                 typecalpres :=  TFmul(Ecran).Q.FindField('PYP_TYPECALPRES').AsString;
                 ExecuteSQL('DELETE FROM PRESENCESALARIE WHERE PYP_SALARIE = "'+salarie+'" AND '+
                             ' PYP_DATEDEBUTPRES = "'+usdatetime(datedebpres)+'" AND ' +
                             ' PYP_DATEFINPRES = "'+usdatetime(datefinpres)+'" AND '+
                             ' PYP_TYPECALPRES = "'+typecalpres+'" AND'+
                             ' PYP_COMPTEURPRES = "'+compteur+'" AND '+
                             ' PYP_PGINDICATPRES <> "AIN" AND PYP_PGINDICATPRES <> "INP"');
                    TFmul(Ecran).Q.Next;
               End;
          End
          // Une partie des salariés a été sélectionnée
          Else
          Begin
               For i := 0 to TFmul(Ecran).Fliste.NbSelected - 1 Do
               Begin

                 {$IFDEF EAGLCLIENT}
                 TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
                 {$ENDIF}
                 TFMul(Ecran).Fliste.GotoLeBOOKMARK(i);
                 Salarie := TFmul(Ecran).Q.FindField('PYP_SALARIE').AsString;
                 datedebpres := TFmul(Ecran).Q.FindField('PYP_DATEDEBUTPRES').AsDatetime;
                 datefinpres := TFmul(Ecran).Q.FindField('PYP_DATEFINPRES').AsDatetime;
                 compteur := TFmul(Ecran).Q.FindField('PYP_COMPTEURPRES').AsString;
                 typecalpres :=  TFmul(Ecran).Q.FindField('PYP_TYPECALPRES').AsString;
                 ExecuteSQL('DELETE FROM PRESENCESALARIE WHERE PYP_SALARIE = "'+salarie+'" AND '+
                             ' PYP_DATEDEBUTPRES = "'+usdatetime(datedebpres)+'" AND ' +
                             ' PYP_DATEFINPRES = "'+usdatetime(datefinpres)+'" AND '+
                             ' PYP_TYPECALPRES = "'+typecalpres+'" AND'+
                             ' PYP_COMPTEURPRES = "'+compteur+'" AND '+
                             ' PYP_PGINDICATPRES <> "AIN" AND PYP_PGINDICATPRES <> "INP"');

               End;
          End;


          TFMul(Ecran).FListe.ClearSelected;
          TFMul(Ecran).BCherche.Click;
     End;
end;

{ TOF_PGANALYSECOMPTEURS }

procedure TOF_PGANALYSECOMPTEURS.OnArgument(S: string);
var Num: Integer;
begin
  inherited;

     // Recherche des libellés des zones TRAVAIL, CODESTAT, LIBREPCMB
  for Num := 1 to 4 do
  begin
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PYP_TRAVAILN' + IntToStr(Num)), GetControl('TPYP_TRAVAILN' + IntToStr(Num)));
  end;

  VisibiliteStat(GetControl('PYP_CODESTAT'), GetControl('TPYP_CODESTAT'));

  for Num := 1 to 4 do
  begin
    VisibiliteChamplibresal(Inttostr(Num), GetControl('PYP_LIBREPCMB' + IntToStr(Num)), GetControl('TPYP_LIBREPCMB' + IntToStr(Num)));
  end;
end;

initialization
  registerclasses([TOF_PGCOMPTEURSCALC_MUL, TOF_PGANALYSECOMPTEURS]);
end.

