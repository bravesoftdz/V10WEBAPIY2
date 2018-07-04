{***********UNITE*************************************************
Auteur  ...... : JP
Créé le ...... : 08/11/2002
Modifié le ... :
Description .. : Source TOF de la FICHE : AFTABBORDOLE
Mots clefs ... : OLE;EXCEL
*****************************************************************}
Unit UTOFAFTABBORDOLE;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     windows,
     messages,
     ent1,
     HTB97,
{$IFDEF EAGLCLIENT}
Maineagl,
{$ELSE}
       db,  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
     forms,M3FP,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HPanel,
     HMsgBox,
     HSysMenu,
     UTOF, utob, vierge,
     SaisUtil, dicobtp, FactUtil, utilCutoff,
     uafo_ressource,
     UTofAfBaseCodeAffaire,
     UtofAFMODIFCUTOFFADD ;

Type
  TOF_AFTABBORDOLE = Class (TOF_AFBASECODEAFFAIRE)
                     procedure OnNew                    ; override ;
                     procedure OnDelete                 ; override ;
                     procedure OnUpdate                 ; override ;
                     procedure OnLoad                   ; override ;
                     procedure OnArgument (S : String ) ; override ;
                     procedure OnClose                  ; override ;

    private
                     GS          :THGrid;
                     TOBOle      :TOB;
                     m_strChamp  :string;

                     procedure AfficheGrid (strColonnes:string);
                     procedure OnExportClick (Sender:TObject);
                     procedure OnImprimeClick (Sender:TObject);

    public
                     Action      :TActionFiche ;
                     EnErreur    :boolean;
                     GModeEclat  :T_ModeEclat;
    end;

procedure AFLanceFiche_TabBordOLE (Argument:String);

Implementation

uses
{$IFNDEF EAGLCLIENT}
mul, printdbg,
{$ELSE}
eMul,
{$ENDIF}
dialogs,
   CbpMCD
   ,CbpEnumerator

;

procedure TOF_AFTABBORDOLE.OnArgument (S:string);
Var
    Critere   :string;
    i, iTableIndex           :integer;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
     inherited;

    MCD := TMCD.GetMcd;
    if not mcd.loaded then mcd.WaitLoaded();

     // La grille de visualisation de la TOB Tableau de bord Ole
     GS := THGRID (GetControl('HGDETAIL'));

     // Récupération des critères
     Critere := (Trim (ReadTokenSt (S)));
     while (Critere <>'') do
     begin
          // TOB tableau de bord OLE
          if Copy (Critere, 1, 6) = 'TOBOLE' then
             TOBOle := TOB (StrToInt (Copy (Critere, 8, Length (Critere)-7)));

          // Champ demandé par le lien OLE
          if Copy (Critere, 1, 8) = 'CHAMPOLE' then
             m_strChamp := Copy (Critere, 10, Length (Critere)-9);

          // Paramètre suivant
          Critere := (Trim (ReadTokenSt (S)));
     end;

     // Titre colonnes
     GS.ColCount    := 9;
     GS.Col         := 0;
     GS.Cells [0,0] := 'Date';
     GS.Cells [1,0] := TraduitGA ('Partie 1 du code affaire');
     GS.Cells [2,0] := TraduitGA ('Partie 2 du code affaire');
     GS.Cells [3,0] := TraduitGA ('Partie 3 du code affaire');
     GS.Cells [4,0] := TraduitGA ('Tiers');
     GS.Cells [5,0] := TraduitGA ('Ressource');
     GS.Cells [6,0] := TraduitGA ('Article');
     GS.Cells [7,0] := 'Libellé';
    table := Mcd.getTable('AFTABLEAUBORD');
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
     begin
          if (FieldList.Current as IFieldCOM).name = 'ATB_AFFAIRE1' then
             GS.Cells [1,0] := (FieldList.Current as IFieldCOM).libelle;
          if (FieldList.Current as IFieldCOM).name = 'ATB_AFFAIRE2' then
             GS.Cells [2,0] := (FieldList.Current as IFieldCOM).libelle;
          if (FieldList.Current as IFieldCOM).name = 'ATB_AFFAIRE3' then
             GS.Cells [3,0] := (FieldList.Current as IFieldCOM).libelle;
          if m_strChamp = (FieldList.Current as IFieldCOM).name then
             GS.Cells [8,0] := (FieldList.Current as IFieldCOM).libelle;
     end;

     // Justification des colonnes
     GS.ColAligns [0] := taCenter;
     GS.ColAligns [1] := taLeftJustify;
     GS.ColAligns [2] := taLeftJustify;
     GS.ColAligns [3] := taLeftJustify;
     GS.ColAligns [4] := taLeftJustify;
     GS.ColAligns [5] := taLeftJustify;
     GS.ColAligns [6] := taLeftJustify;
     GS.ColAligns [7] := taLeftJustify;
     GS.ColAligns [8] := taRightJustify;

     // Sur bouton "export"
     TToolBarButton97 (GetControl ('BEXPORT')).OnClick := OnExportClick;
     TToolBarButton97 (GetControl ('BIMPRIMER')).OnClick := OnImprimeClick;

     // Affichage de la TOB dans la grille
     AfficheGrid ('ATB_DATE;ATB_AFFAIRE1;ATB_AFFAIRE2;ATB_AFFAIRE3;ATB_TIERS;ATB_RESSOURCE;ATB_CODEARTICLE;ATB_LIBELLE');
end;

procedure TOF_AFTABBORDOLE.OnExportClick (Sender:TObject);
var
   i             :integer;
   TOBVisu       :TOB;
   TOBOleFille   :TOB;
   TOBVisuFille  :TOB;
   FileDlg       :TSaveDialog;
   bZero         :boolean;
begin
     // Choix du fichier d'export
     FileDlg := TSaveDialog.Create (Ecran);
     FileDlg.Filter := 'Fichiers texte (*.txt)|*.TXT';
     if FileDlg.Execute = TRUE then
     begin
          // Filtre éventuel sur les lignes à afficher (valeur nulles, ...)
          TOBVisu := TOB.Create ('Détail tableau de bord filtré', nil, -1);
          bZero   := TCheckBox (GetControl ('CBZERO')).Checked;
          for i := 0 to TOBOle.Detail.Count - 1 do
          begin
               TOBOleFille := TOBOle.Detail [i];
               if (bZero = FALSE) or (TOBOleFille.GetValue (m_strChamp) <> 0.0) then
               begin
                    TOBVisuFille := TOB.Create ('AFTABLEAUBORD', TOBVisu, -1);
                    TOBVisuFille.Dupliquer (TOBOleFille, FALSE, TRUE, FALSE);
               end;
          end;

          // Export dans le fichier choisi
          TOBVisu.SaveToFile (FileDlg.FileName, TRUE, TRUE, TRUE);
          TOBVisu.Free;
     end;
     FileDlg.Free;
end;

procedure TOF_AFTABBORDOLE.OnImprimeClick (Sender:TObject);
begin
{$IFNDEF EAGL}
     PrintDBGrid (GS, nil, Ecran.Caption, '');
{$ENDIF}
end;

procedure TOF_AFTABBORDOLE.AfficheGrid (strColonnes:string);
var
   dTotal        :double;
   i             :integer;
   bZero         :boolean;
   TOBVisu       :TOB;
   TOBOleFille   :TOB;
   TOBVisuFille  :TOB;
begin
     TOBVisu := TOB.Create ('Détail tableau de bord filtré', nil, -1);
     bZero   := TCheckBox (GetControl ('CBZERO')).Checked;
     dTotal  := 0.0;

     // Filtre éventuel sur les lignes à afficher (valeur nulles, ...)
     for i := 0 to TOBOle.Detail.Count - 1 do
     begin
          TOBOleFille := TOBOle.Detail [i];
          if (bZero = FALSE) or (TOBOleFille.GetValue (m_strChamp) <> 0.0) then
          begin
               TOBVisuFille := TOB.Create ('AFTABLEAUBORD', TOBVisu, -1);
               TOBVisuFille.Dupliquer (TOBOleFille, FALSE, TRUE, FALSE);
               dTotal := dTotal + TOBOleFille.GetValue (m_strChamp);
          end;
     end;
     TOBVisu.PutGridDetail (GS, FALSE, FALSE, strColonnes + ';' + m_strChamp, TRUE);
     TFVierge(Ecran).Hmtrad.ResizeGridColumns (GS);
     THNumEdit (GetControl ('NETOTAL')).Value := dTotal;

     TOBVisu.Free;
end;

procedure TOF_AFTABBORDOLE.OnNew;
begin
     inherited ;
end;

procedure TOF_AFTABBORDOLE.OnDelete;
begin
     inherited ;
end;

procedure TOF_AFTABBORDOLE.OnUpdate;
begin
end;

procedure TOF_AFTABBORDOLE.OnClose;
begin
     inherited;
end;

procedure TOF_AFTABBORDOLE.OnLoad;
begin
     inherited;
end;

procedure AFLanceFiche_TabBordOLE (Argument:String);
begin
     V_PGI.ZoomOLE := True;   //pour passer la fiche en modal
     AGLLanceFiche ('AFF','AFTABBORDOLE', '', '', Argument);
     V_PGI.ZoomOLE := False;   //pour passer la fiche en modal
end;

procedure AGLDetailTabBordOle (parms:array of variant; nb:integer);
var
     F      :TForm ;
     TOTOF  :TOF;
begin
      F := TForm (Longint(Parms[0])) ;
      if (F is TFVierge) then
      begin
          TOTOF := TFVierge (F).LaTOF;
          if (TOTOF is TOF_AFTABBORDOLE) then
             TOF_AFTABBORDOLE (TOTOF).AfficheGrid ('ATB_DATE;ATB_AFFAIRE1;ATB_AFFAIRE2;ATB_AFFAIRE3;ATB_TIERS;ATB_RESSOURCE;ATB_CODEARTICLE;ATB_LIBELLE');
      end;
end;




Initialization
              registerclasses ( [ TOF_AFTABBORDOLE ] ) ;
              registerAglProc ('AGLDetailTabBordOle', TRUE, 0, AGLDetailTabBordOle);
end.
