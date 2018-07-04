unit AlerteAgenda;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Buttons, ComCtrls, StdCtrls, utob, contnrs, TrayIcon, extctrls, Hctrls;

type
  TFAlerteAgenda = class(TForm)
    LBEnCours: TListBox;
    BFermer: TButton;
    BNOALERT: TButton;
    LEnCours: TLabel;
    LRetard: TLabel;
    LBRetard: TListBox;
    SBNoAlerteRetard: THSpeedButton;
    SBFinRetard: THSpeedButton;
    SBNoAlerteEnCours: THSpeedButton;
    SBReportAlerteEnCours: THSpeedButton;
    TIAgenda: TTrayIcon;
    hcbDelaiReport: THValComboBox;
    LDelai: TLabel;
    procedure LBEnCoursDblClick(Sender: TObject);
    procedure BFermerClick(Sender: TObject);
    procedure LBRetardDblClick(Sender: TObject);
    procedure BNOALERTClick(Sender: TObject);
    procedure LBEnCoursDrawItem (Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
    procedure LBRetardDrawItem (Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
    procedure SBNoAlerteRetardClick(Sender: TObject);
    procedure SBFinRetardClick(Sender: TObject);
    procedure SBNoAlerteEnCoursClick(Sender: TObject);
    procedure SBReportAlerteEnCoursClick(Sender: TObject);

    procedure FormActivate   (Sender: TObject);
    procedure FormDeactivate (Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TIAgendaLeftBtnDblClick(Sender: TObject);

    procedure EnableModif;
    procedure DisableModif;
    procedure LBRetardKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure LBEnCoursKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCreate(Sender: TObject);

  protected
           m_TOBAlerte          :TOB;
           m_Timer              :TTimer;
           m_bHide              :boolean;
           m_oldHintHidePause   :integer;
           m_strDroit           :string;

           procedure DrawListBox     (LBox:TListBox; iIndex:integer; Rect:TRect);
           procedure SetTobAlerte    (TOBAlerte:TOB);
           procedure GetSelectEvt    (SelListe:TObjectList; bRetard:boolean);

           procedure BuildAlerteList;
           procedure AfficheAgenda   (TOBAgenda:TOB);

           function  ActiviteModifiable (TOBAgenda:TOB):boolean;
           procedure TermineActivite; //    (bRetard:boolean);
           procedure SuppressionAlerte  (bRetard:boolean);
           procedure ReportAlerte; //       (bRetard:boolean);

  public
        property TOBAlerte         :TOB        read m_TOBAlerte write SetTobAlerte;
        property Timer             :TTimer                      write m_Timer;
  end;

procedure ShowAlerteAgenda (AgendaAlerte:TOB; AgendaTimer:TTimer);


implementation

{$R *.DFM}

uses hmsgbox, DpOutilsAgenda, galAgenda, //utofagenda, { $$$ todo: transfert GI activiteutil, }
     hent1, dbtables, paramsoc, entdp;

procedure ShowAlerteAgenda (AgendaAlerte:TOB; AgendaTimer:TTimer);
begin
     with TFAlerteAgenda.Create (Application) do
     begin
          TOBAlerte := AgendaAlerte;
          Timer     := AgendaTimer;
          Show;
     end;
end;

procedure TFAlerteAgenda.BuildAlerteList;
var
   i           :integer;
   TOBAgenda   :TOB;
   strItem     :string;
begin
     LBRetard.Clear;
     LBEnCours.Clear;
     for i := 0 to TOBAlerte.Detail.Count-1 do
     begin
          TOBAgenda := TOBAlerte.Detail [i];

          // On signale les retards
          strItem := FormatDateTime ('"le" dd/mm/yyyy "à" hh"h"nn', TOBAgenda.GetValue ('EVTDATE')) + ' - ' + TOBAgenda.GetValue ('EVTCODE') + ': ' + TOBAgenda.GetValue ('EVTLIB');
          if Now > TOBAgenda.GetValue ('EVTDATE') then
              LBRetard.Items.AddObject (strItem, TOBAgenda)
          else
              LBEnCours.Items.AddObject (strItem, TOBAgenda);
     end;

     // Si aucun élément, on grise les boutons
     if LBRetard.Items.Count = 0 then
     begin
          SBNoAlerteRetard.Enabled := FALSE;
          SBFinRetard.Enabled      := FALSE;
     end
     else
     begin
          SBNoAlerteRetard.Enabled := TRUE;
          SBFinRetard.Enabled      := TRUE;
     end;
     if LBEnCours.Items.Count = 0 then
     begin
          SBNoAlerteEnCours.Enabled     := FALSE;
          SBReportAlerteEnCours.Enabled := FALSE;
          hcbDelaiReport.Enabled        := FALSE;
          lDelai.Enabled                := FALSE;
     end
     else
     begin
          SBNoAlerteEnCours.Enabled     := TRUE;
          SBReportAlerteEnCours.Enabled := TRUE;
          hcbDelaiReport.Enabled        := TRUE;
          lDelai.Enabled                := TRUE;
     end;
end;

procedure TFAlerteAgenda.GetSelectEvt (SelListe:TObjectList;bRetard:boolean);
var
   LBSel       :TListBox;
   i           :integer;
   TOBAgenda   :TOB;
begin
     // Sur quel type d'activité: celle en retard ou à venir?
     if bRetard = TRUE then
         LBSel := LBRetard
     else
         LBSel := LBEnCours;

     // Identification des lignes sélectionnées
     for i := 0 to LBSel.Items.Count-1 do
     begin
          if LBSel.Selected [i] = TRUE then
          begin
               TOBAgenda := TOB (LBSel.Items.Objects [i]);
               if TOBAgenda <> nil then
                  SelListe.Add (TOBAgenda);
          end;
     end;
end;

procedure TFAlerteAgenda.SetTobAlerte (TOBAlerte:TOB);
begin
     m_TOBAlerte := TOBAlerte;
     m_bHide     := FALSE;

     BuildAlerteList;
end;

function TFAlerteAgenda.ActiviteModifiable (TOBAgenda:TOB):boolean;
var
   bAbsence   :boolean;
begin
     bAbsence := TOBAgenda.GetValue ('EVTABSENCE') = 'X';
     Result := (m_strDroit = '4') OR ((m_strDroit = '3') and (bAbsence = FALSE)) OR ((m_strDroit = '2') and (bAbsence = TRUE));
end;

procedure TFAlerteAgenda.AfficheAgenda (TOBAgenda:TOB);
var
   strGuidEvt   :string; // $$$ JP 15/03/06 - strNoEvt     :string;
   TOBMaj       :TOB;
begin
     // Affichage en consultation ou en modification si l'utilisateur connecté a les droits
     strGuidEvt := TOBAgenda.GetValue ('JEV_GUIDEVT'); // $$$ JP 15/03/06 - strNoEvt := IntToStr (TOBAgenda.GetValue ('JEV_NOEVT'));
     if ActiviteModifiable (TOBAgenda) then
     begin
          if AgendaLanceFiche (strGuidEvt{// $$$ JP 15/03/06 - strNoEvt}, 'ACTION=MODIFICATION;JEV_EXTERNE=' + TOBAgenda.GetValue ('EVTEXTERNE') + ';JEV_ABSENCE=' + TOBAgenda.GetValue ('EVTABSENCE')) <> '' then
          begin
               TOBMaj := TOB.Create ('maj evt', nil, -1);
               try
                  TOBMaj.LoadDetailFromSQL ('SELECT JEV_GUIDEVT,JEV_DATE AS EVTDATE,JEV_EVTLIBELLE AS EVTLIBELLE,' + // $$$ JP 15/03/06 -  JEV_NOEVT, JEV_DATE AS EVTDATE, JEV_EVTLIBELLE AS EVTLIB, ' +
                                            'JEV_CODEEVT AS EVTCODE,JEV_EXTERNE AS EVTEXTERNE,JEV_ABSENCE AS EVTABSENCE,' +
                                            'JEV_FAIT,JEV_ALERTE FROM JUEVENEMENT WHERE JEV_GUIDEVT="' + strGuidEvt + '"'); // $$$ JP 15/03/06 - JEV_NOEVT=' + strNoEvt);
                  if TOBMaj.Detail.Count > 0 then
                  begin
                       // Il ne faut pas conserver une alerte désactivée ou une activité réalisée
                       if (TOBMaj.Detail [0].GetValue ('JEV_FAIT') = 'X') OR (TOBMaj.Detail [0].GetValue ('JEV_ALERTE') = '-') then
                           TOBAgenda.Free
                       else
                           TOBAgenda.Dupliquer (TOBMaj.Detail [0], FALSE, TRUE);
                       TOBMaj.Free;
                  end
                  else
                      // Il ne faut pas conserver une activité supprimée
                      TOBAgenda.Free;
               except
                     PgiInfo ('La modification ne peut pas être reprise dans l''alerte agenda');
                     TOBMaj.Free;
               end;

               // Mise à jour des listes d'alerte
               BuildAlerteList;
          end;
     end
     else
         AgendaLanceFiche (strGuidEvt {// $$$ JP 15/03/06 - strNoEvt}, 'ACTION=CONSULTATION');
end;

procedure TFAlerteAgenda.LBEnCoursDblClick(Sender: TObject);
var
   TOBAgenda    :TOB;
begin
     TOBAgenda := TOB (LBEnCours.Items.Objects [LBEnCours.ItemIndex]);
     if TOBAgenda <> nil then
        AfficheAgenda (TOBAgenda);
end;

procedure TFAlerteAgenda.LBRetardDblClick(Sender: TObject);
var
   TOBAgenda    :TOB;
begin
     TOBAgenda := TOB (LBRetard.Items.Objects [LBRetard.ItemIndex]);
     if TOBAgenda <> nil then
        AfficheAgenda (TOBAgenda);
end;

procedure TFAlerteAgenda.TermineActivite; // (bRetard:boolean); //BTermineeClick(Sender: TObject);
var
   i          :integer;
   strMessage :string;
   strGuidEvt :string; // $$$ JP 15/03/06 - strNoEvt   :string;
   SelListe   :TObjectList;
begin
     strMessage := 'Confirmez-vous la réalisation des activités sélectionnées ?';
     if PgiAsk (strMessage) = mrYes then
     begin
          SelListe := TObjectList.Create (TRUE);
          GetSelectEvt (SelListe, TRUE);
          if SelListe.Count = 0 then
              PgiInfo ('Veuillez sélectionner au moins une activité')
          else
          begin
               strGuidEvt := ''; // $$$ JP 15/03/06 - strNoEvt := '';
               for i := 0 to SelListe.Count-1 do
               begin
                    if i > 0 then
                       strGuidEvt := strGuidEvt + '","'; // $$$ JP 15/03/06 - strNoEvt := strNoEvt + ',';
                    strGuidEvt := strGuidEvt + TOB (SelListe [i]).GetString ('JEV_GUIDEVT'); // $$$ JP 15/03/06 - strNoEvt := strNoEvt + IntToStr (TOB (SelListe [i]).GetValue ('JEV_NOEVT'));
               end;

               // $$$ JP 29/03/07: si génération en gi désactivée, il faut bien faire la màj du champ "FAIT" quand même
               if GetParamSocSecur ('SO_AGEGENEREGI', True) = TRUE then
                   AgendaSendToGi ('"' + strGuidEvt + '"', TRUE)
               else
                   ExecuteSQL ('UPDATE JUEVENEMENT SET JEV_FAIT="X" WHERE JEV_FAIT<>"X" AND JEV_GUIDEVT IN ("' + strGuidEvt + '")');
          end;

          SelListe.Free;
          BuildAlerteList;
     end;
end;

procedure TFAlerteAgenda.ReportAlerte; // (bRetard:boolean); //Click(Sender: TObject);
var
   i          :integer;
   strGuidEvt :string; // $$$ JP 15/03/06 - strNoEvt   :string;
   strMessage :string;
   SelListe   :TObjectList;
   dtReport   :TDateTime;
   strDecal   :string;
begin
     // Calcul de la date de report: date courante + délai choisi
     strDecal := hcbDelaiReport.Value;
     if strDecal <> '' then
     begin
          if strDecal [1] = '+' then
              // En nb de minutes
              dtReport := Now + StrToTime ('00:' + Copy (strDecal, 2, 2))
          else
              // En nb d'heure
              dtReport := Now + (StrToInt (strDecal) / 24);
     end
     else
     begin
          PgiInfo ('Veuillez définir un délai de report');
          exit;
     end;

     strMessage := 'Confirmez-vous le report des alertes sélectionnées au ' + FormatDateTime ('dd/mm/yyyy "à" hh:nn', dtReport) + ' ?'; //d''alerte des ';
     if PgiAsk (strMessage) = mrYes then // + ' ?') = mrYes then
     begin
          SelListe := TObjectList.Create (TRUE);
          GetSelectEvt (SelListe, FALSE); //bRetard);

          if SelListe.Count = 0 then
              PgiInfo ('Veuillez sélectionner au moins une activité')
          else
          begin
               strGuidEvt := ''; // $$$ JP 15/03/06 - strNoEvt := '';
               for i := 0 to SelListe.Count-1 do
               begin
                    if i > 0 then
                       strGuidEvt := strGuidEvt + '","'; // $$$ JP 15/03/06 - strNoEvt := strNoEvt + ',';
                    strGuidEvt := strGuidEvt + TOB(SelListe [i]).GetString ('JEV_GUIDEVT'); // $$$ JP 15/03/06 - strNoEvt := strNoEvt + IntToStr (TOB (SelListe [i]).GetValue ('JEV_NOEVT'));
               end;
               try
                  // $$$ JP 20/10/06: màj date modification (important pour synchro)
                  ExecuteSQL ('UPDATE JUEVENEMENT SET JEV_ALERTEDATE="' + FormatDateTime ('mm/dd/yyyy hh:nn', dtReport) + '",JEV_DATEMODIF="' + USDATETIME (Now) + '" WHERE JEV_GUIDEVT IN ("' + strGuidEvt + '") AND JEV_ALERTE="X"');
               except
                     PgiInfo ('Impossible de mettre à jour les activités sélectionnées');
               end;
          end;

          SelListe.Free;
          BuildAlerteList;
     end;
end;

procedure TFAlerteAgenda.SuppressionAlerte (bRetard:boolean);
var
   i          :integer;
   strGuidEvt :string; // $$$ JP 15/03/06 - strNoEvt   :string;
   strMessage :string;
   SelListe   :TObjectList;
begin
     strMessage := 'Confirmez-vous la désactivation de l''alerte des ';
     if bRetard = TRUE then
         strMessage := strMessage + 'retards sélectionnés'
     else
         strMessage := strMessage + 'activités sélectionnées';
     if PgiAsk (strMessage + ' ?') = mrYes then
     begin
          SelListe := TObjectList.Create (TRUE);
          GetSelectEvt (SelListe, bRetard);

          if SelListe.Count = 0 then
              PgiInfo ('Veuillez sélectionner au moins une activités')
          else
          begin
               strGuidEvt := ''; // $$$ JP 15/03/06 - strNoEvt := '';
               for i := 0 to SelListe.Count-1 do
               begin
                    if i > 0 then
                       strGuidEvt := strGuidEvt + '","'; // $$$ JP 15/03/06 - strNoEvt := strNoEvt + ',';
                    strGuidEvt := strGuidEvt + TOB (SelListe [i]).GetString ('JEV_GUIDEVT'); // $$$ JP 15/03/06 - strNoEvt := strNoEvt + IntToStr (TOB (SelListe [i]).GetValue ('JEV_NOEVT'));
               end;
               try
                  // $$$ JP 20/10/06: màj date modification (important pour synchro)
                  ExecuteSQL ('UPDATE JUEVENEMENT SET JEV_ALERTE="-",JEV_DATEMODIF="' + USDATETIME (Now) + '" WHERE JEV_GUIDEVT IN ("' + strGuidEvt + '")');
               except
                     PgiInfo ('Impossible de mettre à jour les activités sélectionnées');
               end;
          end;

          SelListe.Free;
          BuildAlerteList;
     end;
end;

procedure TFAlerteAgenda.BFermerClick (Sender: TObject);
begin
     m_bHide := FALSE;
     Close;
end;

procedure TFAlerteAgenda.BNOALERTClick(Sender: TObject);
begin
     if PgiAsk ('Confirmez-vous la désactivation (pour cette session) des alertes de votre agenda ?') = mrYes then
     begin
          m_bHide := TRUE;
          Close;
     end;
end;

procedure TFAlerteAgenda.DrawListBox (LBox:TListBox; iIndex:integer; Rect:TRect);
var
   TOBAgenda    :TOB;
begin
     TOBAgenda := TOB (LBox.Items.Objects [iIndex]);
     if TOBAgenda <> nil then
     begin
          // Affichage du texte
          LBox.Canvas.FillRect (Rect);
          LBox.Canvas.TextOut (Rect.Left+20, Rect.Top, LBox.Items [iIndex]);

          // Dessin d'un rectangle de la couleur définie pour les 3 types d'activité
          if TOBAgenda.GetValue ('EVTABSENCE') = 'X' then
              LBox.Canvas.Brush.Color := rgbAbsence
          else
              if TOBAgenda.GetValue ('EVTEXTERNE') = 'X' then
                  LBox.Canvas.Brush.Color := rgbExterne
              else
                  LBox.Canvas.Brush.Color := rgbInterne;
          Rect.Left   := Rect.Left + 2;
          Rect.Right  := Rect.Left + 10;
          Rect.Top    := Rect.Top + 1;
          Rect.Bottom := Rect.Bottom - 2;
          LBox.Canvas.FillRect (Rect);
     end;
end;

procedure TFAlerteAgenda.LBEnCoursDrawItem (Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
begin
     DrawListBox (TListBox (Control), Index, Rect);
end;

procedure TFAlerteAgenda.LBRetardDrawItem (Control: TWinControl;Index: Integer; Rect: TRect; State: TOwnerDrawState);
begin
     DrawListBox (TListBox (Control), Index, Rect);
end;

procedure TFAlerteAgenda.SBNoAlerteRetardClick (Sender: TObject);
begin
     SuppressionAlerte (TRUE);
end;

procedure TFAlerteAgenda.SBFinRetardClick(Sender: TObject);
begin
     TermineActivite; // (TRUE);
end;

procedure TFAlerteAgenda.SBNoAlerteEnCoursClick(Sender: TObject);
begin
     SuppressionAlerte (FALSE);
end;

procedure TFAlerteAgenda.SBReportAlerteEnCoursClick(Sender: TObject);
begin
     ReportAlerte; // (FALSE);
end;

procedure TFAlerteAgenda.FormActivate(Sender: TObject);
var
   TOBDroit     :TOB;
   i            :integer;
begin
     if Not VH_DP.EnvUserOuvert then exit;
     
     m_oldHintHidePause := Application.HintHidePause;
     Application.HintHidePause := 5000;

     // Chargement droit utilisateur connecté sur lui-même, pour permettre modification ou non sur les activité "zoomées"
     m_strDroit := '0';
     TOBDroit := TOB.Create ('les droits', nil, -1);
     try
        // $$$ JP 16/08/06: nouvelle clé YX_CODE (avec séparateur #)
        TOBDroit.LoadDetailFromSQL ('SELECT YX_LIBRE FROM CHOIXEXT WHERE YX_TYPE="DAU" AND YX_CODE="' + V_PGI.User + '#' + V_PGI.User + '"');
        if TOBDroit.Detail.Count > 0 then
           m_strDroit := TOBDroit.Detail [0].GetValue ('YX_LIBRE');
     finally
            TOBDroit.Free;
     end;

     // Si une fiche agenda est ouverte, on ne permet aucun traitement: report alerte, ...
     for i := 0 to Screen.Formcount-1 do
         if Screen.Forms[i].Name = 'YYAGENDA_FIC' then
            DisableModif;
end;

procedure TFAlerteAgenda.FormDeactivate(Sender: TObject);
begin
     Application.HintHidePause := m_oldHintHidePause;
end;

procedure TFAlerteAgenda.FormShow(Sender: TObject);
begin
     // $$$ JP 01/06/04: avertissement sonore
     if m_TOBAlerte.Detail.Count > 0 then
        Beep;
end;

procedure TFAlerteAgenda.FormClose(Sender: TObject; var Action: TCloseAction);
begin
     if (m_bHide = FALSE) and (VH_DP.EnvUserOuvert) then
        m_Timer.Enabled := TRUE;
     TOBAlerte.Free;
     Action := caFree;
end;

procedure TFAlerteAgenda.TIAgendaLeftBtnDblClick(Sender: TObject);
begin
     Application.Restore;
     Application.BringToFront;
     Self.Show;
end;


procedure TFAlerteAgenda.EnableModif;
begin
     SBNoAlerteRetard.Enabled      := TRUE;
     SBFinRetard.Enabled           := TRUE;
     SBNoAlerteEnCours.Enabled     := TRUE;
     SBReportAlerteEnCours.Enabled := TRUE;
     LBEnCours.Enabled             := TRUE;
     LBRetard.Enabled              := TRUE;
end;

procedure TFAlerteAgenda.DisableModif;
begin
     SBNoAlerteRetard.Enabled      := FALSE;
     SBFinRetard.Enabled           := FALSE;
     SBNoAlerteEnCours.Enabled     := FALSE;
     SBReportAlerteEnCours.Enabled := FALSE;
     LBEnCours.Enabled             := FALSE;
     LBRetard.Enabled              := FALSE;
end;

procedure TFAlerteAgenda.LBRetardKeyDown (Sender:TObject; var Key:Word; Shift:TShiftState);
var
   i     :integer;
   bSel  :boolean;
begin
     bSel := TRUE;
     i    := 0;
     with LBRetard do
     begin
          // Sélection ou désélection ?
          while (bSel = TRUE) and (i < Items.Count) do
          begin
               if Selected [i] = TRUE then
                  bSel := FALSE;
               i := i + 1;
          end;

          // Applique la (dé)-sélection sur tout les éléments
          if (Key = vk_SelectAll) and (Shift = [ssCtrl]) then
             for i := 0 to Items.Count-1 do
                 Selected [i] := bSel;
     end;
end;

procedure TFAlerteAgenda.LBEnCoursKeyDown (Sender:TObject; var Key:Word; Shift:TShiftState);
var
   i     :integer;
   bSel  :boolean;
begin
     bSel := TRUE;
     i    := 0;
     with LBEnCours do
     begin
          // Sélection ou désélection ?
          while (bSel = TRUE) and (i < Items.Count) do
          begin
               if Selected [i] = TRUE then
                  bSel := FALSE;
               i := i + 1;
          end;

          // Applique la (dé)-sélection sur tout les éléments
          if (Key = vk_SelectAll) and (Shift = [ssCtrl]) then
             for i := 0 to Items.Count-1 do
                 Selected [i] := bSel;
     end;
end;

procedure TFAlerteAgenda.FormCreate (Sender:TObject);
begin
     hcbDelaiReport.ItemIndex := 1;
end;



end.

