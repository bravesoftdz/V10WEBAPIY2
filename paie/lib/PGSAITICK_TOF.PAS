{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 14/05/    
Modifié le ... : 06/06/2003
Description .. : Source TOF de la FICHE : PGSAITICK ()
Suite ........ : saisie de la commande de tickets restaurant
Mots clefs ... : TOF;PGSAITICK
*****************************************************************}
{
 PT1    MF    10/02/2005  V_6.0  1-Il ne faut pas initialiser PSD_LIBELLE.
                                 Car lors du calcul du bulletin une ligne de
                                 commentaire est générée si PSD_LIBELLE <> ''.
                                 2- modif. traitement de recup du nbre de tickets
                                 (NbTicket = partie entière de la valuer DOUBLE
                                 de la cellule)
 PT2    MF    14/02/2005  V_6.0  Correction requête car les champs de la liste
                                 ne sont pas les mêmes que l'on soit en saisie
                                 ou en intégration.
 PT3    MF    16/02/2005  V_6.0  Correction erreur qd valeur faciale = 0
 PT4    MF    25/02/2005  V_6.0  Les modification lors de la saisie étaient
                                 enregistrées à tort (même si non validée)
 PT5    MF    06/12/2006  V_7.0  Correction formattage valeur faciale (de type double) pour alimentation
                                 tob de maj.
 PT6    NA    23/06/2008  V_850  Chèque déjeuner : le code client est soit celui du paramsoc, coit celui saisi
                                 par établissement
}
Unit PGSAITICK_TOF ;

Interface

Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
//unused     StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     PrintDbg,
{$ELSE}
     UtileAgl,
     HQry,
{$ENDIF}
     dialogs,
//unused     forms,
     sysutils,
//unused     ComCtrls,
     HCtrls,
     HEnt1,
     HXlsPas,
     HMsgBox,
     UTOF,
     vierge,
     ed_tools,
     UTOB,
     P5Util,
     HTB97,
     ParamSoc,
     Entpaie, // pt6
     Grids;

Type
  TOF_PGSAITICK = Class (TOF)
    private
    QMul                                        : TQUERY;     // Query recuperee du mul
    LaGrille                                    : THGrid;     // Grille de saisie
    Modifier, BoolIntegr, BoolCde,SauvSais      : Boolean;
    DateDebut,DateFin                           : TDateTime;
    TSal                                        : TOB;
    TypeTraitement                              : String;

    procedure GrilleCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
    procedure ValiderClick(Sender: TObject);
    procedure FermeClick(Sender: TObject);
    function  OnSauve: boolean;
    procedure MajTOBT_Sais(var T : TOB; TSalFille : TOB; CodeClient : string; FraisGestion, NbTickets : double; NoOrdre,i : Integer);
    procedure MajTOBT_Mvt(var T : TOB; TSalFille,LeTypeTicket : TOB; CodeClient : string; NbTickets : double);
    procedure ImpClik (Sender: TObject);
    procedure BtnExportClick(Sender: TObject);
    procedure BtnDeleteClick(Sender: TObject);

    public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_PGSAITICK.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGSAITICK.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGSAITICK.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGSAITICK.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 06/06/2003
Modifié le ... : 06/06/2003
Description .. : Procédure OnArgument
Suite ........ : Si le paramètre société SO_PGTYPEPREPACDE = "001"
Suite ........ : on charge les éléments du mois précédent (table CDETICKETS)
Suite ........ : sinon les salariés sont présentés en fonction du type de 
Suite ........ : tickets paramétré.
Suite ........ : Chargement de la grille de saisie
Suite ........ : Préparation de la TOB (TSal) qui sera  utilisée pour la maj de 
Suite ........ : CDETICKETS
Mots clefs ... : TOF;PGSAITICK
*****************************************************************}
procedure TOF_PGSAITICK.OnArgument (S : String ) ;
var
  F                                                     : TFVierge ;
  Ligne,i                                               : Integer ;
  okok,Abandon                                          : boolean ;
  St, TypePrepaCde                                      : string;
  Q, Qtt                                                : TQuery;
  TTYPETICKET,LeTypeTicket,TSalFille,TBulletin          : TOB;
  CdeMoisPrec,CdeMoisEnCours,Cde, TBullFille            : TOB;
  BtnValid, BtnFerme, BtnExport,BtnDelete               : TToolbarButton97;
  BtnImp                                                : TToolbarButton97;
  PPatronale                                            : double;
  DateDMoisPrec,DateFMoisPrec                           : TDateTime;
  CdeExistante, CdeE                                    : TOB;
begin
  Inherited ;

  TSal := NIL;
  TTYPETICKET := NIL;
  Modifier := False;
  PPatronale := 0.0;
  Abandon := FALSE;

  TypePrepaCde := GetParamSocSecur('SO_PGTYPEPREPACDE','');  

  TSal := TOB.Create ('Les Salaries Tickets',NIL, -1);

  St := Trim (S);
  DateDebut := StrToDate(ReadTokenSt(st));  // Recup de la date de debut
  DateFin := StrToDate (ReadTokenSt(st));   // Recup Date fin
  TypeTraitement := (ReadTokenSt(st));      // Récup type de traitement (S= Saisie Cde, I= Intégration)
  DateDMoisPrec := PlusMois(DateDebut,-1);  // Calcul date de debut mois précédent
  DateFMoisPrec := FindeMois(PlusMois(DateFin,-1)); // Calcul date de fin mois précédent

  F:=TFVierge(Ecran) ;
  if F <> NIL then
  begin
{$IFDEF EAGLCLIENT}
   QMUL:=THQuery(F.FMULQ).TQ ;
{$ELSE}
   QMUL:=F.FMULQ ;
{$ENDIF}
  BtnImp    := TToolbarButton97 (GetControl ('Bimprimer'));
  if BtnImp <> NIL then BtnImp.OnClick := ImpClik;
  SetControlEnabled('Bimprimer',True);
  SetControlVisible('Bimprimer',True);

  BtnValid:=TToolbarButton97 (GetControl ('BVALIDER'));
  if BtnValid<>NIL then BtnValid.OnClick := ValiderClick ;
  BtnFerme:=TToolbarButton97 (GetControl ('BFERME'));
  if BtnFerme<>NIL then BtnFerme.OnClick := FermeClick ;
  BtnExport := TToolbarButton97(GetControl('BEXPORTER'));
  if BtnExport<>nil then BtnExport.OnClick:=BtnExportClick;
  BtnDelete:= TToolbarButton97(GetControl('BDELETE'));
  if BtnDelete<>nil then BtnDelete.OnClick:=BtnDeleteClick;
  if (TypeTraitement =  'C') then
    {il existe déjà une cde pendant la période à d'autres dates
     la suppression est possible}
    BtnDelete.Visible := True
  else
    BtnDelete.Visible := False;
  end ;
  if QMUL = NIL then exit;
  if (TypeTraitement = 'S') then
    {saise de la commande}
    Ecran.Caption := Ecran.Caption + '  '+
                     RechDom ('TTUTILISATEUR', V_PGI.User, FALSE)+
                     '  du '+DateToStr(DateDebut)+
                     ' au '+DateToStr(DateFin)+' ('+
                     RechDom ('PGTYPEPREPACDE',TypePrepaCde, FALSE)+')'
  else
    if (TypeTraitement = 'I') then
      {saisie de modification avant intrégration}
      Ecran.Caption := 'Intégration dans la paie' + '  '+
                     RechDom ('TTUTILISATEUR', V_PGI.User, FALSE)+
                     '  du '+DateToStr(DateDebut)+
                     ' au '+DateToStr(DateFin)
    else
      if (TypeTraitement = 'C') then
        {consulation de cde déjà existante pendant la période}
        Ecran.Caption := 'déjà saisie par' + '  '+
                     RechDom ('TTUTILISATEUR', V_PGI.User, FALSE)+
                     ' pour la période  du '+DateToStr(QMul.FindField ('PRT_DATEDEBUT').AsDateTime)+
                     ' au '+DateToStr(QMul.FindField ('PRT_DATEFIN').AsDateTime);

  if (TypeTraitement <> 'C') then
  begin
    {Confection de la TOB CdeMoisEnCours pour travailler en modification sur
    commandes déjà saisies}
    CdeMoisEnCours := TOB.Create ('La commande mois en cours',NIL,-1);
    St := 'SELECT PRT_SALARIE, PRT_NBTICKET '+
          'FROM CDETICKETS '+
          'WHERE PRT_DATEDEBUT = "'+USDateTime(DateDebut)+'" '+
          'AND PRT_DATEFIN = "'+USDateTime(DateFin)+'"';
    Q := OpenSql (St, TRUE);
    if not Q.EOF then
      CdeMoisEnCours.LoadDetailDB ('COMMANDE TICKETS M', '','',Q,FALSE,FALSE);
   Ferme(Q);

   if (TypePrepaCde = '001') then
   { on repart de la commande du mois précédent }
   begin
     CdeMoisPrec := TOB.Create ('La commande mois précédent',NIL,-1);
     St := 'SELECT PRT_SALARIE, PRT_NBTICKET '+
           'FROM CDETICKETS '+
           'WHERE PRT_DATEDEBUT = "'+USDateTime(DateDMoisPrec)+'" '+
           'AND PRT_DATEFIN = "'+USDateTime(DateFMoisPrec)+'"';
     Q := OpenSql (St, TRUE);
     if not Q.EOF then
       CdeMoisPrec.LoadDetailDB ('COMMANDE TICKETS M-1', '','',Q,FALSE,FALSE);
     Ferme(Q);
    end;
  end;
  LaGrille := THGrid(Getcontrol('GRILLE'));
  if LaGrille <> nil then
  begin
    LaGrille.OnCellExit  := GrilleCellexit;

    LaGrille.ColWidths[2] := 120;
    LaGrille.ColWidths[3] := 120;
    LaGrille.ColWidths[4] := 200;
    LaGrille.ColWidths[5] := 50;
    LaGrille.ColWidths[6] := 80;
    LaGrille.ColAligns[0] := taCenter;
    LaGrille.ColAligns[5] := taRightJustify;
    LaGrille.ColAligns[6] := taRightJustify;

//    if (TypeTraitement = 'C') then
      {La colonne nbre de tickets n'est pas modifiable}
//      LaGrille.FixedCols := 7;


    i := QMUL.RecordCount ;
    if (TypeTraitement <> 'C') then
      {traitement de saisie}
      InitMoveProgressForm (NIL,'Chargement des données de la saisie',
                            'Veuillez patienter SVP ...',i,TRUE,TRUE);
    if (TypeTraitement <> 'C') then
    begin
      TTypeTicket := TOB.Create ('Les type de tickets restaurant',NIL,-1);
      St := 'SELECT PTT_TYPTICKET, PTT_NBTICKET, PTT_VFACIALE, PTT_PPATRONALE  '+
            'FROM TICKETRESTAURANT '+
            'ORDER BY PTT_TYPTICKET';
      Q := OpenSql (St, TRUE);
      TTYPETICKET.LoadDetailDB ('TICKET', '','',Q,FALSE,FALSE);
      Ferme(Q);
    end;
    if (TypeTraitement <> 'C') then
    {Traitement de saisie}
    begin
      if (TypePrepaCde = '004') then
      {Cde à partir des infos bulletin}
      begin
        QMul.First ;
        TBulletin := TOB.Create('Les tickets du bulletin',NIL,-1);
        While Not QMul.EOF do
        begin
          st := 'SELECT PHB_SALARIE,PHB_BASEREM FROM TICKETRESTAURANT LEFT JOIN HISTOBULLETIN '+
                'ON PHB_RUBRIQUE =  PTT_RETSALARIAL ';
//d PT2
          if (TypeTraitement = 'S') then
            st := st+'WHERE PHB_SALARIE = "'+QMul.FindField ('PSA_SALARIE').AsString+'"'
          else
            st := st+'WHERE PHB_SALARIE = "'+QMul.FindField ('PRT_SALARIE').AsString+'"';
          st := st+' AND ';
          if (TypeTraitement = 'S') then
            st := st+'PTT_TYPTICKET = "'+ QMul.FindField ('PSE_TYPTICKET').AsString+'"'
          else
            st := st+'PTT_TYPTICKET = "'+ QMul.FindField ('PRT_TYPTICKET').AsString+'"';
// PT2
          st := st+' AND '+
                'PHB_DATEDEBUT = "'+USDateTime(DateDebut)+'" '+
                ' AND '+
                'PHB_DATEFIN = "'+USDateTime(DateFin)+'"';
          Q := OpenSql (St, TRUE);
          if not Q.EOF then
          begin
            TBullFille := TOB.Create('',TBulletin,-1);
            TBullFille.AddChampSupValeur('SALARIE',Q.FindField ('PHB_SALARIE').AsString,
                                         FALSE);
            TBullFille.AddChampSupValeur('NBTICKET',
                                         StrToInt(FormatFloat('0',Q.FindField ('PHB_BASEREM').AsFloat)),
                                         FALSE);
          end;
          Ferme (Q);
          QMul.Next;
        end;
      end;

      QMul.First ;
      LaGrille.RowCount := 2;
      Ligne := 1; // Ligne courante
      Abandon := FALSE;
      While Not QMul.EOF do
      begin
        if (TypeTraitement = 'S') then
          {Traitement de saisie avant envoi}
          LeTypeTicket := TTYPETICKET.FindFirst (['PTT_TYPTICKET'],
                                                 [QMul.FindField ('PSE_TYPTICKET').AsString],
                                                 TRUE)
        else
          {Traitement de saisie avant intégration dans la paie}
          LeTypeTicket := TTYPETICKET.FindFirst (['PTT_TYPTICKET'],
                                                 [QMul.FindField ('PRT_TYPTICKET').AsString],
                                                 TRUE);

        if LeTypeTicket <> NIL then
        begin
          { Code distribution}
          if (TypeTraitement = 'S') then
           {Traitement de saisie avant envoi}
            LaGrille.Cells [0,Ligne] := QMul.FindField ('PSE_DISTRIBUTION').AsString
          else
            {Traitement de saisie avant intégration dans la paie}
            LaGrille.Cells [0,Ligne] := QMul.FindField ('PRT_DISTRIBUTION').AsString;

          { Matricule salarié}
          if (TypeTraitement = 'S') then
            {Traitement de saisie avant envoi}
            LaGrille.Cells [1,Ligne] := QMul.FindField ('PSA_SALARIE').AsString
          else
            {Traitement de saisie avant intégration dans la paie}
            LaGrille.Cells [1,Ligne] := QMul.FindField ('PRT_SALARIE').AsString;

          { Nom Préniom}
          LaGrille.Cells [2,Ligne] := Trim(QMul.FindField ('PSA_LIBELLE').AsString)+
                                      ' '+
                                      Trim(QMul.FindField ('PSA_PRENOM').AsString) ;
          { Personnalisation}
          if (TypeTraitement = 'S') then
            {Traitement de saisie avant envoi}
            LaGrille.Cells [3,Ligne] := copy(QMul.FindField ('PSE_PERSONNAL').AsString,1,14)
          else
            {Traitement de saisie avant intégration dans la paie}
            LaGrille.Cells [3,Ligne] := copy(QMul.FindField ('PRT_PERSONNAL').AsString,1,14);

          { Information complémentaire }
          if (TypeTraitement = 'S') then
            {Traitement de saisie avant envoi}
            LaGrille.Cells [4,Ligne] := QMul.FindField ('PSE_INFOCOMPL').AsString
          else
            {Traitement de saisie avant intégration dans la paie}
            LaGrille.Cells [4,Ligne] := QMul.FindField ('PRT_INFOCOMPL').AsString;

          { Valeur faciale}
          LaGrille.Cells [5,Ligne] := DoubleToCell(LeTypeTicket.GetValue ('PTT_VFACIALE'),2);

          { Nombre de tickets}
          if (TypePrepaCde = '003') then
          {Préparation automatique en fonction du type de ticket paramétré}
          begin
            Cde := CdeMoisEnCours.FindFirst(['PRT_SALARIE'],
                                            [LaGrille.Cells [1,Ligne]],
                                            TRUE);
            if Cde <> NIL then
              {il y a déjà une commande pour ce salarié pour cette période}
              LaGrille.Cells [6,Ligne] := DoubleToCell(Cde.GetValue('PRT_NBTICKET'),0)
            else
              LaGrille.Cells [6,Ligne] := DoubleToCell(LeTypeTicket.GetValue ('PTT_NBTICKET'),0);
          end;

          if (TypePrepaCde = '001') then
          { Reprise de la quantité du mois précédent}
          begin
            Cde := CdeMoisEnCours.FindFirst(['PRT_SALARIE'],
                                            [LaGrille.Cells [1,Ligne]],
                                            TRUE);
            if Cde <> NIL then
              {il y a déjà une commande pour ce salarié pour cette période}
              LaGrille.Cells [6,Ligne] := DoubleToCell(Cde.GetValue('PRT_NBTICKET'),0)
            else
            begin
              Cde := CdeMoisPrec.FindFirst(['PRT_SALARIE'],
                                           [LaGrille.Cells [1,Ligne]],
                                           TRUE);
              if Cde <> NIL then
               { Cde du mois précédent}
               LaGrille.Cells [6,Ligne] := DoubleToCell(Cde.GetValue('PRT_NBTICKET'),0)
              else
               { selon type de ticket}
               LaGrille.Cells [6,Ligne] := DoubleToCell(LeTypeTicket.GetValue('PTT_NBTICKET'),0);
            end;
          end;

          if (TypePrepaCde = '004') then
          {Reprise de la quantité du bulletin en cours}
          begin
            Cde := CdeMoisEnCours.FindFirst(['PRT_SALARIE'],
                                            [LaGrille.Cells [1,Ligne]],
                                            TRUE);
            if Cde <> NIL then
              {il y a déjà une commande pour ce salarié pour cette période}
              LaGrille.Cells [6,Ligne] := DoubleToCell(Cde.GetValue('PRT_NBTICKET'),0)
            else
            begin
              Cde := TBulletin.FindFirst(['SALARIE'],
                                         [LaGrille.Cells [1,Ligne]],
                                         TRUE);
              if Cde <> NIL then
               {Nbre de tickets récupéré sur bulletin de la période}
               LaGrille.Cells [6,Ligne] := DoubleToCell(Cde.GetValue('NBTICKET'),0)
              else
               {Selon type de ticket}
               LaGrille.Cells [6,Ligne] := DoubleToCell(LeTypeTicket.GetValue('PTT_NBTICKET'),0);
            end;
          end;

          if (LaGrille.Cells [6,Ligne] = '') then
            LaGrille.Cells [6,Ligne] := '0';

          PPatronale := LeTypeTicket.GetValue ('PTT_PPATRONALE');

          Modifier := TRUE ;
          Ligne := Ligne + 1;
          LaGrille.RowCount := LaGrille.RowCount + 1;
        end;
        okok := MoveCurProgressForm (IntToStr(Ligne)) ;
        if Not OkOk then
        begin
          Abandon := TRUE;
          break;
        end;

        if LeTypeTicket <> NIL then
        begin
          TSalFille := TOB.Create('',TSal,-1);
          if (TypeTraitement = 'S') then
         {Traitement de saisie avant envoi}
          begin
            TSalFille.AddChampSupValeur('ETABLISSEMENT',
                                        QMul.FindField ('PSA_ETABLISSEMENT').AsString,
                                        FALSE);
            TSalFille.AddChampSupValeur('MATRICULE',
                                        QMul.FindField ('PSA_SALARIE').AsString,
                                        FALSE);
          end
          else
         {Traitement de saisie avant intégration dans la paie}
          begin
            TSalFille.AddChampSupValeur('ETABLISSEMENT',
                                        QMul.FindField ('PRT_ETABLISSEMENT').AsString,
                                        FALSE);
            TSalFille.AddChampSupValeur('MATRICULE',
                                        QMul.FindField ('PRT_SALARIE').AsString,
                                        FALSE);
          end;
          if (Typetraitement = 'S') then
            {Traitement de saisie avant envoi}
            TSalFille.AddChampSupValeur('TYPTICKET',
                                        QMul.FindField ('PSE_TYPTICKET').AsString,
                                        FALSE)
          else
            {Traitement de saisie avant intégration dans la paie}
            TSalFille.AddChampSupValeur('TYPTICKET',
                                        QMul.FindField ('PRT_TYPTICKET').AsString,
                                        FALSE);
          if (Typetraitement = 'I') then
          begin
            {Traitement de saisie avant intégration dans la paie}
            TSalFille.addChampSupValeur('NBTICKET',
                                        QMul.FindField ('PRT_NBTICKET').AsInteger,
                                        FALSE);
            TSalFille.addChampSupValeur('CONFIDENTIEL',
                                        QMul.FindField ('PSA_CONFIDENTIEL').AsInteger,
                                        FALSE);
          end;

          TSalFille.AddChampSupValeur('PPATRONALE',PPatronale);
          TSalFille.AddChampSupValeur('NOMPRENOM',
                                      copy(
                                      Trim(QMul.FindField ('PSA_LIBELLE').AsString)+
                                      ' '+
                                      Trim(QMul.FindField ('PSA_PRENOM').AsString),
                                      1,20),
                                      FALSE);
          if (TypeTraitement = 'S') then
          {Traitement de saisie avant envoi}
          begin
            TSalFille.AddChampSupValeur('CODESTAT',
                                        QMul.FindField ('PSA_CODESTAT').AsString,
                                        FALSE);
            TSalFille.AddChampSupValeur('TRAVAILN1',
                                        QMul.FindField ('PSA_TRAVAILN1').AsString,
                                        FALSE);
            TSalFille.AddChampSupValeur('TRAVAILN2',
                                        QMul.FindField ('PSA_TRAVAILN2').AsString,
                                        FALSE);
            TSalFille.AddChampSupValeur('TRAVAILN3',
                                        QMul.FindField ('PSA_TRAVAILN3').AsString,
                                        FALSE);
            TSalFille.AddChampSupValeur('TRAVAILN4',
                                        QMul.FindField ('PSA_TRAVAILN4').AsString,
                                        FALSE);
            TSalFille.AddChampSupValeur('LIBRE1','',FALSE);
            TSalFille.AddChampSupValeur('LIBRE2','',FALSE);
            TSalFille.AddChampSupValeur('LIBRE3','',FALSE);
            TSalFille.AddChampSupValeur('LIBRE4','',FALSE);
            TSalFille.AddChampSupValeur('BOOLIBRE1','',FALSE);
            TSalFille.AddChampSupValeur('BOOLIBRE2','',FALSE);
            TSalFille.AddChampSupValeur('BOOLIBRE3','',FALSE);
            TSalFille.AddChampSupValeur('BOOLIBRE4','',FALSE);
          end
          else
          {Traitement de saisie avant intégration dans la paie}
          begin
            TSalFille.AddChampSupValeur('CODESTAT',
                                        QMul.FindField ('PRT_CODESTAT').AsString,
                                        FALSE);
            TSalFille.AddChampSupValeur('TRAVAILN1',
                                        QMul.FindField ('PRT_TRAVAILN1').AsString,
                                        FALSE);
            TSalFille.AddChampSupValeur('TRAVAILN2',
                                        QMul.FindField ('PRT_TRAVAILN2').AsString,
                                        FALSE);
            TSalFille.AddChampSupValeur('TRAVAILN3',
                                        QMul.FindField ('PRT_TRAVAILN3').AsString,
                                        FALSE);
            TSalFille.AddChampSupValeur('TRAVAILN4',
                                        QMul.FindField ('PRT_TRAVAILN4').AsString,
                                        FALSE);
            TSalFille.AddChampSupValeur('LIBRE1','',FALSE);
            TSalFille.AddChampSupValeur('LIBRE2','',FALSE);
            TSalFille.AddChampSupValeur('LIBRE3','',FALSE);
            TSalFille.AddChampSupValeur('LIBRE4','',FALSE);
            TSalFille.AddChampSupValeur('BOOLIBRE1','',FALSE);
            TSalFille.AddChampSupValeur('BOOLIBRE2','',FALSE);
            TSalFille.AddChampSupValeur('BOOLIBRE3','',FALSE);
            TSalFille.AddChampSupValeur('BOOLIBRE4','',FALSE);
          end;
        end;

        QMul.NEXT;
      end;
    end;
    if (TypeTraitement = 'C') then
    { cas de commande déjà existante pendant la période
    { affichage des lignes pour suppression éventuelle}
    begin
      TTypeTicket := TOB.Create ('Les type de tickets restaurant',NIL,-1);
      St := 'SELECT PTT_TYPTICKET, PTT_VFACIALE '+
            'FROM TICKETRESTAURANT ';
      Qtt := OpenSql (St, TRUE);
      TTYPETICKET.LoadDetailDB ('TICKET', '','',Qtt,FALSE,FALSE);

//      QMul.First ;

      St := 'SELECT PRT_SALARIE, PRT_NBTICKET, PRT_DISTRIBUTION, '+
            'PRT_PERSONNAL, PRT_INFOCOMPL,PRT_TYPTICKET, '+
            'PRT_DATEINTEG, PRT_DATECDE '+
            'FROM CDETICKETS '+
            'WHERE PRT_DATEDEBUT = "'+
            USDateTime(QMul.FindField ('PRT_DATEDEBUT').AsDateTime)+'" '+
            'AND PRT_DATEFIN = "'+
            USDateTime(QMul.FindField ('PRT_DATEFIN').AsDateTime)+'"';

      Q := OpenSql (St, TRUE);
      if not Q.EOF then
       CdeExistante := TOB.Create ('La commande existante',NIL,-1);
       CdeExistante.LoadDetailDB ('COMMANDE TICKETS E', '','',Q,FALSE,FALSE);
      Ferme(Q);
      CdeE := CdeExistante.FindFirst([''],[''], TRUE);
      Ligne := 1;
      LaGrille.RowCount := 2;
      BoolCde := False;
      BoolIntegr := False;
      While  CdeE <> NIL do
      begin
        LeTypeTicket := TTYPETICKET.FindFirst (['PTT_TYPTICKET'],
                                               [CdeE.GetValue('PRT_TYPTICKET')],
                                               TRUE);
        LaGrille.Cells [0,Ligne] := CdeE.GetValue ('PRT_DISTRIBUTION');
        LaGrille.Cells [1,Ligne] := CdeE.GetValue('PRT_SALARIE');
        St := 'SELECT PSA_LIBELLE, PSA_PRENOM FROM SALARIES '+
              'WHERE PSA_SALARIE = "'+CdeE.GetValue('PRT_SALARIE')+'"';
        Q := OpenSql (St, TRUE);
        if not Q.EOF then
        LaGrille.Cells [2,Ligne] := Trim(Q.FindField ('PSA_LIBELLE').AsString)+
                                      ' '+
                                      Trim(Q.FindField ('PSA_PRENOM').AsString) ;
        ferme(Q);

        LaGrille.Cells [3,Ligne] := CdeE.GetValue('PRT_PERSONNAL');
        LaGrille.Cells [4,Ligne] := CdeE.GetValue('PRT_INFOCOMPL');

        if LeTypeTicket <> NIL then
          LaGrille.Cells [5,Ligne] := DoubleToCell(LeTypeTicket.GetValue ('PTT_VFACIALE'),2);
        LaGrille.Cells [6,Ligne] := DoubleToCell(CdeE.GetValue('PRT_NBTICKET'),0);
        if (LaGrille.Cells [6,Ligne] = '') then
          LaGrille.Cells [6,Ligne] := '0';

        if (CdeE.GetValue('PRT_DATEINTEG') <> IDate1900) then
          BoolIntegr := True;
        if (CdeE.GetValue('PRT_DATECDE') <> IDate1900) then
          BoolCde := True;
        CdeE := CdeExistante.FindNext([''],[''], TRUE);
        Ligne := Ligne + 1;
        LaGrille.RowCount := LaGrille.RowCount + 1;
      end;
      Ferme(Qtt);
    end;
  end;
  FiniMoveProgressForm();
  if Abandon = TRUE then Close; // Abandon de la saisie

  LaGrille.RowCount := LaGrille.RowCount - 1;

  FreeAndNil(TTYPETICKET);

  FreeAndNil(CdeMoisPrec);
  FreeAndNil(CdeMoisEnCours);
  FreeAndNil(CdeExistante);
// if V_PGI.Debug=True then
//      PGVisuUnObjet( TSal,'','');
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 06/06/2003
Modifié le ... : 06/06/2003
Description .. : procédure OnClose
Suite ........ : sortie de la fiche : on libère les TOB
Mots clefs ... :  TOF;PGSAITICK
*****************************************************************}
procedure TOF_PGSAITICK.OnClose ;
begin
  Inherited ;
  FreeAndNil(TSal);
end ;

procedure TOF_PGSAITICK.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGSAITICK.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 06/06/2003
Modifié le ... :   /  /    
Description .. : procédure GrilleCellExit
Suite ........ : sortie d'une cellule . vérif la valeur saisie est numérique et 
Suite ........ : supérieure ou égale à zéro
Mots clefs ... : TOF;PGSAITICK
*****************************************************************}
procedure TOF_PGSAITICK.GrilleCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  Modifier := TRUE;
  if (LaGrille.Row-1 = 0) AND (ARow = 1) then exit;
  if not IsNumeric(LaGrille.Cells[Acol, Arow]) then
    Lagrille.Row := LaGrille.Row-1;
  if Valeur(LaGrille.Cells[Acol, Arow]) < 0 then
    Lagrille.Row := LaGrille.Row-1;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 06/06/2003
Modifié le ... :   /  /    
Description .. : procédure ValiserClick
Suite ........ : demande de validation
Mots clefs ... : TOF;PGSAITICK
*****************************************************************}
procedure TOF_PGSAITICK.ValiderClick(Sender: TObject);
var
  rep : Integer;
begin
inherited ;
  if (TypeTraitement = 'C') then
  {Traitement de consultation}
  exit;

  if (Modifier = True) then
  begin
    rep:=PGIAsk ('Voulez vous sauvegarder votre saisie ?', Ecran.Caption) ;
    if rep=mrNo then
    begin
      SauvSais := False;
      OnSauve;
    end;
    if rep=mrCancel then exit;
    if rep=mryes then
    begin
      SauvSais := True;
      OnSauve;
    end;

  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 06/06/2003
Modifié le ... : 06/06/2003
Description .. : procédure FermerClick
Suite ........ : sur action de la croix on lance la prpcédure de valisation et 
Suite ........ : on quiite la fiche
Mots clefs ... : TOF;PGSAITICK
*****************************************************************}
procedure TOF_PGSAITICK.FermeClick(Sender: TObject);
begin
  if (Modifier = True) then ValiderClick(Sender);
  Close;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 06/06/2003
Modifié le ... : 06/06/2003
Description .. : fonction OnSauve
Suite ........ : A partir de la tob Tsal et de la grille alimentation de la tob 
Suite ........ : Tsais qui est utilisée pour la maj de la table CDETICKETS
Mots clefs ... : TOF;PGSAITICK
*****************************************************************}
function TOF_PGSAITICK.OnSauve: boolean;
var
    st                                                          : String;
    i, rep                                                      : integer;
    T_Sais,T,TSalFille,TTypeTicket,TOBMvtTick,LeTypeTicket      : TOB;
    CodeClient                                                  : string;
    FraisGestion                                                : double;
    NoOrdre,NbTickets                                           : integer;
    Q                                                           : TQuery;

begin
  T_Sais := NIL;

  result := TRUE;
  if Modifier = FALSE then exit;

  CodeClient := GetParamSocSecur('SO_PGCODECLIENT','');
  FraisGestion :=  Valeur(GetParamSocSecur('SO_PGFRAISGESTION',0.00000));
  if (TypeTraitement = 'S') then
  {Saisie de la commande de tickets}
  begin
    try
    BeginTrans;
    Modifier  := FALSE;
    T_Sais := TOB.Create ('Les lignes de saisie', NIL, -1);
    TSalFille := TSal.FindFirst([''],[''], TRUE);
    for i := 1 to LaGrille.RowCount - 1 do
    begin
//PT1-2      NbTickets := StrToInt(LaGrille.Cells [6,i]);
      NbTickets := Trunc(Valeur(LaGrille.Cells [6,i]));

      NoOrdre := 0;
      T := TOB.Create ('CDETICKETS', T_Sais, -1);
      if T <> NIL then
      begin
        if sauvsais then  // PT4
          MajTOBT_Sais(T,TSalFille,CodeClient,FraisGestion, NbTickets,NoOrdre,i);
      end;
      TSalFille := TSal.FindNext([''],[''], TRUE);
    end;  // Boucle sur les lignes de la grille
// d PT4
   if sauvsais then
    begin
// f PT4
      T_Sais.SetAllModifie (TRUE);
      T_Sais.InsertOrUpdateDB (TRUE);
    end; // PT4
    CommitTrans;
    except
    result := FALSE;
    Rollback;
    PGIBox ('Une erreur est survenue lors de la validation de la saisie','');
    end;;
  FreeAndNil(T_Sais);
  end;

  if (TypeTraitement = 'I') then
  {Intégration dans la paie}
  begin
    rep:=PGIAsk ('Confirmez-vous l''intégration de cette commande dans la paie?',
                 Ecran.Caption) ;
    if rep=mryes then
    begin
      TTypeTicket := TOB.Create ('Les type de tickets restaurant',NIL,-1);
      St := 'SELECT PTT_TYPTICKET, PTT_RETSALARIAL, PTT_RETPATRONALE '+
            'FROM TICKETRESTAURANT '+
            'ORDER BY PTT_TYPTICKET';
      Q := OpenSql (St, TRUE);
      TTypeTicket.LoadDetailDB ('TICKET', '','',Q,FALSE,FALSE);
      Ferme(Q);

      TOBMvtTick := TOB.Create('Mouvement à intégrer',NIl,-1);
      TSalFille := TSal.FindFirst([''],[''], TRUE);
      for i := 1 to LaGrille.RowCount - 1 do
      begin
        if (SauvSais = True) then
          NbTickets := StrToInt(LaGrille.Cells [6,i])
        else
          NbTickets := TSalFille.Getvalue('NBTICKET');

        T := TOB.Create ('HISTOSAISRUB', TOBMvtTick, -1);
        if T <> NIL then
        begin
{          LeTypeTicket := TTYPETICKET.FindFirst (['PTT_TYPTICKET'],
                                                  [QMul.FindField ('PRT_TYPTICKET').AsString],
                                                  TRUE);}
          LeTypeTicket := TTYPETICKET.FindFirst (['PTT_TYPTICKET'],
                                                 [TSalFille.GetValue('TYPTICKET')],
                                                 TRUE);
          if (LeTypeTicket.GetValue ('PTT_RETSALARIAL') <> '') and
             (LeTypeTicket.GetValue ('PTT_RETSALARIAL') <> NULL) then
          { Retenue SALARIALE }
          begin
            MajTOBT_Mvt(T,TSalFille,LeTypeTicket,CodeClient, NbTickets);
          end;
          if (LeTypeTicket.GetValue ('PTT_RETPATRONALE') <> '') and
             (LeTypeTicket.GetValue ('PTT_RETPATRONALE') <> NULL) then
          { Retenue Patronale }
          begin
            T := TOB.Create ('HISTOSAISRUB', TOBMvtTick, -1);
            MajTOBT_Mvt(T,TSalFille,LeTypeTicket,CodeClient, NbTickets);
          end;
        end;
        TSalFille := TSal.FindNext([''],[''], TRUE);
      end;  {Boucle sur les lignes de la grille}

      try
      begintrans;
      if TOBMvtTick <> NIL then
      begin // PT1
        TOBMvtTick.SetAllModifie (TRUE); // PT1
        TOBMvtTick.InsertOrUpdateDB(FALSE) ;
      end; // PT1
      Committrans;
      except
      Rollback;
      PGIBox('! Erreur maj table HISTOSAISRUB','');
      end;

      try
      BeginTrans ;
      QMul.First;
      While Not QMul.EOF do
      begin
        St := 'UPDATE CDETICKETS SET PRT_DATEINTEG = "'+
              USDateTime(Date)+'" '+
              'WHERE '+
              'PRT_SALARIE = "'+QMul.findfield('PRT_SALARIE').AsString+'" AND '+
              'PRT_DATEDEBUT = "'+USDateTime(QMul.findfield('PRT_DATEDEBUT').AsDateTime)+'" '+
              'AND PRT_DATEFIN = "'+USDateTime(QMul.findfield('PRT_DATEFIN').AsDateTime)+'"';
        ExecuteSQL(St);
        QMul.Next;
      end;
      CommitTrans;
      except
        RollBack;
        PGIBox('! Erreur maj table CDETICKETS','');
      end ;
      PGIInfo('Terminé','Intégration');

      FreeAndNil(TOBMvtTick);
      FreeAndNil(TTypeTicket);
    end;
  end;
  Modifier := FALSE;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 06/06/2003
Modifié le ... : 23/06/2003
Description .. : procédure MajTobTSais
Suite ........ : alimentaion d'un élément de la tob T_sais pour mettre à jour 
Suite ........ : la table CDETICKETS
Mots clefs ... : TOF;PGSAITICK
*****************************************************************}
procedure TOF_PGSAITICK.MajTOBT_Sais(var T : TOB; TSalFille : TOB; CodeClient : string; FraisGestion, NbTickets : double; NoOrdre,i : Integer);
// deb pt6
var
fournisseur, Etabl, st : string;
Q : TQUERY;
// fin pt6
begin
// deb pt6
  Fournisseur := GetParamSocSecur('SO_PGTYPECDETICKET','');
  if fournisseur = '004' then
  begin
   if (Vh_Paie.PGFACTETABL) then
    begin
     Etabl := TSalFille.GetValue('ETABLISSEMENT');
     // recherche du code client selon l'établissement
     st := 'SELECT ETB_TICKLIVR FROM ETABCOMPL WHERE ETB_ETABLISSEMENT = '+Etabl+'';
     Q := OPENSQL(st, true);
     if Not Q.EOF THEN Codeclient := Q.findfield('ETB_TICKLIVR').asstring;
     FERME(Q);
    end;
   end;
// fin pt6

  T.PutValue ('PRT_CODECLIENT',CodeClient );
  T.PutValue ('PRT_ETABLISSEMENT',TSalFille.GetValue('ETABLISSEMENT'));
  T.PutValue ('PRT_DISTRIBUTION',LaGrille.Cells [0,i]);
  T.PutValue ('PRT_SALARIE',LaGrille.Cells [1,i]);
  T.PutValue ('PRT_DATEDEBUT', DateDebut);
  T.PutValue ('PRT_DATEFIN', DateFin);
  T.PutValue ('PRT_DATECDE',IDate1900);
  T.PutValue ('PRT_ORDRE',NoOrdre);
  T.PutValue ('PRT_DATEINTEG',IDate1900);
  T.PutValue ('PRT_DATECREATION',Date);
  T.PutValue ('PRT_CODESTAT',TSalFille.GetValue('CODESTAT'));
  T.PutValue ('PRT_TRAVAILN1',TSalFille.GetValue('TRAVAILN1'));
  T.PutValue ('PRT_TRAVAILN2',TSalFille.GetValue('TRAVAILN2'));
  T.PutValue ('PRT_TRAVAILN3',TSalFille.GetValue('TRAVAILN3'));
  T.PutValue ('PRT_TRAVAILN4',TSalFille.GetValue('TRAVAILN4'));
  T.PutValue ('PRT_LIBRE1',TSalFille.GetValue('LIBRE1'));
  T.PutValue ('PRT_LIBRE2',TSalFille.GetValue('LIBRE2'));
  T.PutValue ('PRT_LIBRE3',TSalFille.GetValue('LIBRE3'));
  T.PutValue ('PRT_LIBRE4',TSalFille.GetValue('LIBRE4'));
  T.PutValue ('PRT_BOOLIBRE1',TSalFille.GetValue('BOOLIBRE1'));
  T.PutValue ('PRT_BOOLIBRE2',TSalFille.GetValue('BOOLIBRE2'));
  T.PutValue ('PRT_BOOLIBRE3',TSalFille.GetValue('BOOLIBRE3'));
  T.PutValue ('PRT_BOOLIBRE4',TSalFille.GetValue('BOOLIBRE4'));
  T.PutValue ('PRT_PERSONNAL',LaGrille.Cells [3,i]);
  T.PutValue ('PRT_INFOCOMPL',LaGrille.Cells [4,i]);
  T.PutValue ('PRT_NBTICKET',NbTickets);
// d PT3
  if (LaGrille.Cells [5,i] <> '') then
// PT5    T.PutValue ('PRT_VFACIALE',LaGrille.Cells [5,i])
    T.PutValue ('PRT_VFACIALE',StrToFloat(LaGrille.Cells [5,i]))
  else
    T.PutValue ('PRT_VFACIALE','0');
// f PT3
  T.PutValue ('PRT_PPATRONALE',TSalFille.GetValue('PPATRONALE'));
  T.PutValue ('PRT_FRAISGESTION',FraisGestion);
  T.PutValue ('PRT_TYPTICKET',TSalFille.GetValue('TYPTICKET'));
  T.PutValue ('PRT_NOMPRENOM',TSalFille.GetValue('NOMPRENOM'))

end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 23/06/2003
Modifié le ... : 23/06/2003
Description .. : procédure MajTobT_Mvt
Suite ........ : alimentation de la tob TobT_Mvt pour mise à jour de 
Suite ........ : HISTOSAISRUB
Mots clefs ... : PAIE ; PGSAITICK
*****************************************************************}
procedure TOF_PGSAITICK.MajTOBT_Mvt(var T : TOB; TSalFille,LeTypeTicket : TOB; CodeClient : string; NbTickets : double);
var
  blanc                                 : String;
begin
   blanc := StringOfChar(' ',35);
   begin
     T.PutValue('PSD_ORIGINEMVT','SRB');
     T.PutValue('PSD_SALARIE',TSalFille.GetValue('MATRICULE'));
     T.PutValue('PSD_DATEDEBUT',DateDebut);
     T.PutValue('PSD_DATEFIN',DateFin);
     T.PutValue('PSD_RUBRIQUE',LeTypeTicket.GetValue ('PTT_RETSALARIAL'));
     T.PutValue('PSD_ORDRE',1);
// PT1-1    T.PutValue('PSD_LIBELLE',copy(blanc,1,35));
     T.PutValue('PSD_LIBELLE','');
     T.PutValue('PSD_RIBSALAIRE',copy(blanc,1,17));
     T.PutValue('PSD_BANQUEEMIS',copy(blanc,1,35));
     T.PutValue('PSD_TOPREGLE','-');
     T.PutValue('PSD_DATEPAIEMENT',IDate1900);
     T.PutValue('PSD_ETABLISSEMENT',
                TSalFille.GetValue('ETABLISSEMENT'));
     T.PutValue('PSD_TYPALIMPAIE','BAS');
     T.PutValue('PSD_BASE',NbTickets);
     T.PutValue('PSD_TAUX',0);
     T.PutValue('PSD_COEFF',0);
     T.PutValue('PSD_MONTANT',0);
     T.PutValue('PSD_DATEINTEGRAT',0);
     T.PutValue('PSD_DATECOMPT',0);
     T.PutValue('PSD_AREPORTER','');
     T.PutValue('PSD_CONFIDENTIEL',TSalFille.GetValue('CONFIDENTIEL'));
     T.PutValue('PSD_TOPCONVERT','-');
   end;
   if (LeTypeTicket.GetValue ('PTT_RETPATRONALE') <> '') and
      (LeTypeTicket.GetValue ('PTT_RETPATRONALE') <> NULL) then
   { Retenue Patronale }
   begin
     T.PutValue('PSD_ORIGINEMVT','SRB');
     T.PutValue('PSD_SALARIE',TSalFille.GetValue('MATRICULE'));
     T.PutValue('PSD_DATEDEBUT',DateDebut);
     T.PutValue('PSD_DATEFIN',Datefin);
     T.PutValue('PSD_RUBRIQUE',
                              LeTypeTicket.GetValue ('PTT_RETPATRONALE'));
     T.PutValue('PSD_ORDRE',1);
// PT1-1    T.PutValue('PSD_LIBELLE',copy(blanc,1,35));
     T.PutValue('PSD_LIBELLE','');
     T.PutValue('PSD_RIBSALAIRE',copy(blanc,1,17));
     T.PutValue('PSD_BANQUEEMIS',copy(blanc,1,35));
     T.PutValue('PSD_TOPREGLE','-');
     T.PutValue('PSD_DATEPAIEMENT',IDate1900);
     T.PutValue('PSD_ETABLISSEMENT',TSalFille.getValue('ETABLISSEMENT'));
     T.PutValue('PSD_TYPALIMPAIE','BAS');
     T.PutValue('PSD_BASE',NbTickets);
     T.PutValue('PSD_TAUX',0);
     T.PutValue('PSD_COEFF',0);
     T.PutValue('PSD_MONTANT',0);
     T.PutValue('PSD_DATEINTEGRAT',0);
     T.PutValue('PSD_DATECOMPT',0);
     T.PutValue('PSD_AREPORTER','');
     T.PutValue('PSD_CONFIDENTIEL',TSalFille.GetValue('CONFIDENTIEL'));
     T.PutValue('PSD_TOPCONVERT','-');
   end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 16/06/2003
Modifié le ... :   /  /
Description .. : Procédure ImpClick
Suite ........ : Lance l'édition du contenu de la saisie de tickets restaurant
Mots clefs ... : PAIE; PGSAITICK
*****************************************************************}
procedure TOF_PGSAITICK.ImpClik (Sender: TObject);
{$IFDEF EAGLCLIENT}
var

  TLaGrille, TFille                             : TOB;
  i                                             : integer;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
  PrintDBGrid ( TCustomGrid(LaGrille) , NIL, Ecran.Caption,'') ;
{$ELSE}
  if (TLaGrille = NIL) then
    TLaGrille := TOB.Create ('grille à éditer', NIL, -1);
  for i := 1 to LaGrille.RowCount - 1 do
  begin
    if TLaGrille <> NIL then
    begin
      TFille := TOB.Create ('grille à éditer', TLaGrille, -1);
      TFille.AddChampSupValeur ('PRT_DISTRIBUTION',LaGrille.Cells [0,i]);
      TFille.AddChampSupValeur ('PRT_SALARIE',LaGrille.Cells [1,i]);
      TFille.AddChampSupValeur ('PRT_NOMPRENOM',LaGrille.Cells [2,i]);
      TFille.AddChampSupValeur ('PRT_PERSONNAL',LaGrille.Cells [3,i]);
      TFille.AddChampSupValeur ('PRT_INFOCOMPL',LaGrille.Cells [4,i]);
      TFille.AddChampSupValeur ('PRT_VFACIALE',Valeur(LaGrille.Cells [5,i]));
      TFille.AddChampSupValeur ('PRT_NBTICKET',StrToInt(LaGrille.Cells [6,i]));
    end;
  end;  // Boucle sur les lignes de la grille
  LanceEtatTOB('E','PAY','CDT',TLaGrille,True,False,False,LaGrille,'','',False);
  FreeAndNil (TLaGrille);
{$ENDIF}
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 16/06/2003
Modifié le ... :   /  /    
Description .. : procédure BtnExportClick
Suite ........ : Lance l'export vers excel de la grille de saisie de la 
Suite ........ : commande de tickets restaurant
Mots clefs ... : PAIE ; PGSAITICK
*****************************************************************}
procedure TOF_PGSAITICK.BtnExportClick(Sender: TObject);
var
  Grille                : ThGrid;
  SD                    : TSaveDialog;
begin
  Grille:=ThGrid(GetControl('GRILLE'));
  if Grille = nil then exit;
  SD:=TSaveDialog.Create(nil);
  SD.Filter:='Fichier Excel (*.xls)|*.xls';
  if SD.Execute then
  begin
  {Fichier Excel (*.xls)|*.xls|  SD.FilterIndex = 1}
  if (SD.FilterIndex=1) and (Pos('.xls',SD.FileName)<1) then
    SD.FileName:=Trim(SD.FileName)+'.xls';
  ExportGrid(Grille,nil,SD.FileName,SD.FilterIndex,True) ;
  end;
  SD.Destroy;
end;
procedure TOF_PGSAITICK.BtnDeleteClick(Sender: TObject);
var
  St                            : string;
  rep                           : integer;
begin
 if (BoolIntegr = True) then
    begin
      rep:=PGIAsk ('Commande déjà intégrée dans la paie, validez-vous sa suppression?', Ecran.Caption) ;
      if rep=mrNo then exit ;
      if rep=mrCancel then exit;
    end;
    if (BoolCde = True) then
    begin
      rep:=PGIAsk ('Les lignes de la commande ont déjà fait l''objet d''un envoi, validez-vous la suppression?', Ecran.Caption) ;
      if rep=mrNo then exit ;
      if rep=mrCancel then exit;
    end;
  rep:=PGIAsk ('Confirmez-vous la suppression de ces lignes de commande?', Ecran.Caption) ;
  if rep=mrNo then exit ;
  if rep=mrCancel then exit;

 try
    BeginTrans;
    St := 'DELETE FROM CDETICKETS WHERE '+
          'PRT_DATEDEBUT = "'+USDateTime(QMul.FindField ('PRT_DATEDEBUT').AsDateTime)+'" '+
          'AND PRT_DATEFIN = "'+USDateTime(QMul.FindField ('PRT_DATEFIN').AsDateTime)+'"';

    ExecuteSQL(St);
    Committrans;
 except
    Rollback;
    PGIBox('! Erreur suppression elt table CDETICKETS','');
 end;
    PGIInfo('Terminé','Suppression ');
end;
Initialization
  registerclasses ( [ TOF_PGSAITICK ] ) ;
end.



