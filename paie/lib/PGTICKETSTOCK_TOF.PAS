{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 26/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGTICKETSTOCK ()
Mots clefs ... : TOF;PGTICKETSTOCK
*****************************************************************}
{
 PT1    MF    14/02/2005  V_6.0  1- Le nom du pré-fichier de commande contient
                                 les dates de prériode (--> autant de pré-fichiers
                                 que de périodes)
                                 2- Mise en place de la confection du fichier
                                 pour NATEXIS
 PT2    MF    22/01/2007  V_720  Nouveau fournisseur ACCOR                                
}
Unit PGTICKETSTOCK_TOF ;

Interface

Uses
//unused     StdCtrls,
//unused     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     UTOB,
{$ENDIF}
     FileCtrl,
     EntPaie,
//unused     forms,
     sysutils,
//unused     ComCtrls,
     HTB97,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     LookUp,
     ParamSoc,
     P5Util,
     PgOutils2;
//unused     PgOutils ;

Type
  TOF_PGTICKETSTOCK = Class (TOF)
    private
    DateDebut,DateFin   : TDateTime;
    LaGrille            : THGrid;
    BtnIns,BtnDel       : TToolBarButton97;
    IndMaj              : Boolean;
    Fournisseur         : string; // PT1-2

    procedure GrilleCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
    procedure GrilleElipsisClick(Sender: TObject);
    procedure GrilleRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GrilleClick(Sender: TObject);
    procedure BTnInsClick(Sender :Tobject);
    procedure BTnDelClick(Sender: TObject);
    procedure OuvrePreFichier(var FPrefichier : TextFile; var Prefichier : string;var CodeRetour : boolean);
    procedure CreatLigneStock(CodeClient : String;var FPrefichier : TextFile);



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
var
  etab                : string; // PT2
procedure TOF_PGTICKETSTOCK.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGTICKETSTOCK.OnDelete ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... : 11/06/2003
Description .. : Procédure OnUpdtae
Suite ........ : Contrôle du nombre de ticket saisis (entre 0 et 99)
Mots clefs ... : PAIE; PGTICKETSTOCK
*****************************************************************}
procedure TOF_PGTICKETSTOCK.OnUpdate ;
var
  i                        : integer;
begin
  Inherited ;
  IndMaj := True;
  for i := 1 to LaGrille.RowCount-1 do
  begin
    if (LaGrille.Cells[3,i] <> '') and
       ((not IsNumeric(LaGrille.Cells[4, i])) or
       (Valeur(LaGrille.Cells[4,i]) <= 0)or
       (Valeur(LaGrille.Cells[4,i]) > 99)) then
    begin
      IndMaj := False;
      LastError := 1 ;
      PGIBox('La quantité doit être supérieure à zéro et inférieure à 100','Mise à jour impossible');
    end;
// d PT2
  if (Fournisseur = '003') then
  begin
    // ACCOR
    etab := GetControlText('PTLIVR');
    if (etab = '') then
    begin
      LastError := 1 ;
      PgiBox('La point de livraison n''est pas renseigné','');
      SetFocusControl('PTLIVR');
    end;
  end;
// f PT2
  end;

end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... :   /  /    
Description .. : Procédure OnLoad
Suite ........ : Paramétrage des cellules de la grille
Mots clefs ... : PAIE; PGTICKETSTOCK
*****************************************************************}
procedure TOF_PGTICKETSTOCK.OnLoad ;
begin
  Inherited ;
  LaGrille.ColAligns[0] := taCenter;
  LaGrille.ColAligns[1] := taRightJustify;
  LaGrille.ColAligns[2] := taRightJustify;
  LaGrille.ColAligns[3] := taLeftJustify;
  LaGrille.ColAligns[4] := taRightJustify;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... :   /  /    
Description .. : Procédure OnArgument
Suite ........ : préparation de la grille
Mots clefs ... : PAIE; PGTICKETSTOCK
*****************************************************************}
procedure TOF_PGTICKETSTOCK.OnArgument (S : String ) ;
var
  St                            : String;
begin
  Inherited ;
  St:= Trim (S);
  DateDebut := StrToDate(ReadTokenSt(st));  // Recup date de debut
  DateFin := StrToDate (ReadTokenSt(st));   // Recup Date fin

  SetControlText('DATEDEB',DateToStr(DateDebut));
  SetControlText('DATEFIN',DateToStr(DateFin));

  LaGrille := THGrid(Getcontrol('GTICKETSTOCK'));
  if LaGrille <> nil then
  begin
//  LaGrille.OnCellExit  := GrilleCellexit;
   LaGrille.OnCellEnter := GrilleCellEnter;
   LaGrille.OnRowEnter := GrilleRowEnter;
//   LaGrille.OnRowExit := GrilleRowExit;
   LaGrille.OnClick := GrilleClick;
   LaGrille.OnElipsisClick := GrilleElipsisClick;
//  LaGrille.PostDrawCell := GrillePostDrawCell;
//  LaGrille.OnDblClick := GrilleDblClick;
  end
  else exit;


  BtnIns := TToolBarButton97(GetControl('BINS_LINE'));
  if BtnIns <> NIL then
  begin
   BtnIns.OnClick := BTnInsClick;
   end;

   BtnDel := TToolBarButton97(GetControl('BDEL_LINE'));
   if BtnDel <> NIL then
   begin
     BtnDel.OnClick := BTnDelClick;
   end;
   //d PT1-2
  Fournisseur := GetParamSocSecur('SO_PGTYPECDETICKET','');
// d PT2
   if (Fournisseur = '003') then
   begin
     // ACCOR
     SetControlVisible('PTLIVR', TRUE);
     SetControlEnabled('PTLIVR', TRUE);
     SetControlVisible('LPTLIVR', TRUE);
   end;
// f PT2
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... : 11/06/2003
Description .. : Procédure OnClose
Suite ........ : Fermeture de la fiche - lance l'ouverture du Pré fichier de 
Suite ........ : stock et l'alimentation de celui-ci
Mots clefs ... : PAIE; PGTICKETSTOCK
*****************************************************************}
procedure TOF_PGTICKETSTOCK.OnClose ;
var
  FPreFichier                           : TextFile;
  CodeClient,Prefichier                 : string;
  CodeRetour                            : boolean;
begin
  Inherited ;
  if (IndMaj = True) then
  begin
    CodeClient := VH_Paie.PgCodeClient;
    OuvrePreFichier(FPrefichier ,Prefichier,CodeRetour);
    CreatLigneStock(CodeClient , FPrefichier);
  end;
end ;

procedure TOF_PGTICKETSTOCK.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGTICKETSTOCK.OnCancel () ;
begin
  Inherited ;
end ;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... :   /  /
Description .. : Procédure GrilleCellEnter
Suite ........ : A l'entrée sur la cellule 3 on place un elipsis
Mots clefs ... : PAIE; PGTICKETSTOCK
*****************************************************************}
procedure TOF_PGTICKETSTOCK.GrilleCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  if (LaGrille.col = 3) then
    LaGrille.ElipsisButton := TRUE
  else
    LaGrille.ElipsisButton := FALSE;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... :   /  /    
Description .. : Procédure GrilleElipsisClick
Suite ........ : un click sur la cellule 3 provoque le choix du type de ticket
Mots clefs ... : PAIE; PGTICKETSTOCK
*****************************************************************}
procedure TOF_PGTICKETSTOCK.GrilleElipsisClick(Sender: TObject);
var
  Q                             : TQuery;
begin
  if (LaGrille.col = 3) then
    begin
      LookUpList (LaGrille,'Type','TICKETRESTAURANT','PTT_TYPTICKET','PTT_LIBELLE','','PTT_TYPTICKET',TRUE,-1);
      Q := OpenSql('SELECT PTT_NBTICKET,PTT_VFACIALE,PTT_PPATRONALE '+
                   ' FROM TICKETRESTAURANT '+
                   'WHERE PTT_TYPTICKET="'+LaGrille.CellValues[LaGrille.col,LaGrille.row]+'" ',
                   True);
     if not Q.eof then
     begin
       LaGrille.CellValues[1,LaGrille.row] := DoubleToCell(Q.FindField('PTT_VFACIALE').AsFloat,2) ;
       LaGrille.CellValues[2,LaGrille.row] := DoubleToCell(Q.FindField('PTT_PPATRONALE').AsFloat,2) ;
       LaGrille.CellValues[4,LaGrille.row] := DoubleToCell(Q.FindField('PTT_NBTICKET').AsInteger,0) ;
     end;
     Ferme(Q);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... :   /  /    
Description .. : Procédure GrilleRowEnter
Suite ........ : A l'entrée sur une ligne de la grille on se positionne sur la 
Suite ........ : cellule 3
Mots clefs ... : PAIE; PGTICKETSTOCK
*****************************************************************}
procedure TOF_PGTICKETSTOCK.GrilleRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
  NoLigne, NoCol                               : integer;
begin
  LaGrille.Cells[0,LaGrille.row] := 'STOCK';
  NoLigne := LaGrille.row ;
  NoCol := 3;
  GrilleCellEnter(Sender,NoCol,NoLigne, Cancel);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... :   /  /    
Description .. : procédure GrilleClick
Suite ........ : un click sur la grille provoque un entrée sur un ligne
Mots clefs ... : PAIE; PGTICKETSTOCK
*****************************************************************}
procedure TOF_PGTICKETSTOCK.GrilleClick(Sender: TObject);
var
   Vcancel                   : boolean;
begin
   GrilleRowEnter(Sender,0,Vcancel, False);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... :   /  /    
Description .. : Procédure BtnInsClick
Suite ........ : Action bouton nouvelle ligne
Mots clefs ... : PAIE; PGTICKETSTOCK
*****************************************************************}
procedure TOF_PGTICKETSTOCK.BTnInsClick(Sender :Tobject);
begin
  Lagrille.InsertRow(1);
  LaGrille.Col := 3;
  LaGrille.Row := 1;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... :   /  /    
Description .. : Procédure BtnDekClick
Suite ........ : Action bouton suppression de ligne
Mots clefs ... : PAIE ; PGTICKETSTOCK
*****************************************************************}
procedure TOF_PGTICKETSTOCK.BTnDelClick(Sender: TObject);
begin
  if (LaGrille.Row <> 0) then
  begin
    LaGrille.DeleteRow(LaGrille.Row) ;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... :   /  /    
Description .. : Procédure OuvrePreFichier
Suite ........ : Ouverture du fichier PSTOCK.st1
Mots clefs ... : PAIE; PGTICKETSTOCK
*****************************************************************}
procedure TOF_PGTICKETSTOCK.OuvrePreFichier(var FPrefichier : TextFile; var Prefichier : string;var CodeRetour : boolean);
var
  Dir                   : string;
  Periode               : string; // PT1-1
begin
  CodeRetour := True;
//d PT1-1
  Periode := Copy(DateToStr(Datedebut),9,2)+
             Copy(DateToStr(Datedebut),4,2)+
             Copy(DateToStr(Datedebut),1,2)+
             Copy(DateToStr(Datefin),9,2)+
             Copy(DateToStr(Datefin),4,2)+
             Copy(DateToStr(Datefin),1,2) ;
//f PT1-1

  Dir :=  Trim(GetParamSocSecur('SO_PGREPERTTICKET',''));    
  if not DirectoryExists(Dir) then
    if not ForceDirectories(Dir) then
    begin
      PGIInfo('Impossible de créer le répertoire de travail '+Dir,
              'Titres restaurant');
      TraceErr.Items.Add ('Impossible de créer le répertoire de travail '+Dir);
      CodeRetour := False;
    end;

  if (CodeRetour = True) then
  begin
//PT1-1    PreFichier := Dir+'\PSTOCK.st1';
    PreFichier := Dir+'\PSTOCK'+Periode+'.st1';
    if (FileExists(PreFichier)) then
      DeleteFile(PChar(Prefichier));

    AssignFile(FPreFichier, PreFichier);
    {$i-} ReWrite(FPreFichier); {$i+}
    if IoResult <> 0 then
      CodeRetour := False;
  end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... : 11/06/2003
Description .. : Procédure CreatLigneStock
Suite ........ : alimentation du fichier PSTOCK.st1 d'enregistrements de
Suite ........ : fabrication uniquement (Type 2)
Suite ........ : code distribution=STOCK
Suite ........ : alumentation nombre de ticket, valeur faciale, participation
Suite ........ : employeur
Mots clefs ... : PAIE; PGTICKETSTOCK
*****************************************************************}
procedure TOF_PGTICKETSTOCK.CreatLigneStock(CodeClient : String;var FPrefichier : TextFile);
var
  i                                             : integer;
  HuitZero,blanc                                : string;
  NbT, NbCarnet, TotTicket                      : integer;
// d PT1-2
  VFaciale, PPatronale                          : string;
  SepDec                                        : Char;
  NbCarnetLigne,NbTicketCarnet,NbTotalLigne     : integer;
  NbCarnetW,NbCarnetTotal, Reste                : integer;
// f PT1-2
// d PT2
  PointdeLivraison                              : string;
  QQ                                            : TQuery;
  CodeDistribution                              : string;
  EdNPCarnet,EdNPTicket,EdRSCarnet,EdCPVCarnet  : string;
  CodePostal, Ville, ANbt                       : string;
  NiveauInterne1,NiveauInterne2,RaisonSociale   : string;
  j                                             : integer;
// f PT2
begin
  TotTicket := 0;
  NbCarnet := 0;

  blanc := StringOfChar(' ',61);
  HuitZero := '00000000';

// d PT2
  if (Fournisseur = '003') then
  // ACCOR
  begin
    // Point de livraison
    if (V_PGI.NumVersionBase < 800) then
    // (spécif V7)
    begin
      QQ := opensql('SELECT ETB_AUTRENUMERO FROM ETABCOMPL WHERE ETB_ETABLISSEMENT = "'+etab+'"',TRUE);
      if not QQ.EOF then
      begin
        PointdeLivraison :=  QQ.findfield('ETB_AUTRENUMERO').asstring;
      end;
      ferme (QQ);
    end
    else
    // (à partir V8)
    begin
      QQ := opensql('SELECT ETB_TICKLIVR FROM ETABCOMPL WHERE ETB_ETABLISSEMENT = "'+etab+'"',TRUE);
      if not QQ.EOF then
      begin
        PointdeLivraison :=  QQ.findfield('ETB_TICKLIVR').asstring;
      end;
      ferme (QQ);
    end;
    // Code distribution
    CodeDistribution := '';
    QQ := opensql('SELECT CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE="PCD" AND CC_ABREGE="STOCK"',TRUE);
    if not QQ.EOF then
    begin
      CodeDistribution :=  QQ.findfield('CC_LIBRE').asstring;
    end;
    ferme (QQ);
    RaisonSociale := GetParamSocSecur('SO_LIBELLE','');
    PgUpperCase(RaisonSociale);
  end;
// f PT2

  for i := 1 to LaGrille.RowCount-1 do
  begin
    NbT := StrToInt(FormatFloat('#00',Valeur(LaGrille.CellValues[4,i])) );
    TotTicket := TotTicket + NbT;
    if ((NbT div 25) * 25) < NbT then
      NbCarnet := NbCarnet + (NbT div 25)+1
    else
      NbCarnet := NbCarnet + (NbT div 25);

// d PT1-2
    if (Fournisseur = '001') then
    begin
    // SODEXHO
// f PT1-2
      Writeln(FPreFichier,copy(HuitZero,1,(8-length(Trim(CodeClient))))+
                       Trim(CodeClient)+
                       'STOCK     '+
                       ColleZeroDevant(StrToInt(LaGrille.CellValues[4,i]),2)+
                       copy(HuitZero,1,(5-length(FormatFloat('#00',Valeur(LaGrille.CellValues[1,i])*100))))+
                       FormatFloat('#00',Valeur(LaGrille.CellValues[1,i])*100)+
                       copy(blanc,1,11)+
                       copy(blanc,1,20)+
                       copy(blanc,1,14)+
                       copy(HuitZero,1,(5-length(FormatFloat('#00',Valeur(LaGrille.CellValues[2,i])*100))))+
                       FormatFloat('#00',Valeur(LaGrille.CellValues[2,i])*100)+
                       '    '+
                       '2');
// d PT1-2
    end
    else
    if (Fournisseur = '002') then // PT2
    begin
    // NATEXIS
      // (5 car avec . décimal)
      SepDec := DecimalSeparator;
      DecimalSeparator := '.';

      if (length(FormatFloat('00.00',Valeur(LaGrille.CellValues[1,i]))) = 5) then
        VFaciale := FormatFloat('00.00',Valeur(LaGrille.CellValues[1,i]))
      else
        VFaciale := copy(HuitZero,1,(5-length(FormatFloat('00.00',Valeur(LaGrille.CellValues[1,i])))))+
                 FormatFloat('00.00',Valeur(LaGrille.CellValues[1,i]));

      if (length(FormatFloat('00.00',Valeur(LaGrille.CellValues[2,i]))) = 5) then
        PPatronale := FormatFloat('00.00',Valeur(LaGrille.CellValues[2,i]))
      else
        PPatronale := copy(HuitZero,1,(5-length(FormatFloat('00.00',Valeur(LaGrille.CellValues[2,i])))))+
                   FormatFloat('00.00',Valeur(LaGrille.CellValues[2,i]));

      DecimalSeparator := SepDec;

      if ((NbT div 25) * 25) < NbT then
         NbCarnetLigne := (NbT div 25)+1
       else
         NbCarnetLigne := (NbT div 25);

      if (NbCarnetLigne > 999) then
      // Le nbre de carnets par ligne de cde est limité à 999
      begin
        NbCarnetW := 999;
        NbCarnetTotal := NbCarnetLigne;
        while (NbCarnetW < NbCarnetTotal) do
        begin
          NbcarnetLigne := 999;
          NbTicketCarnet := 25;
          NbTotalLigne := NbCarnetLigne * NbTicketCarnet;
          Writeln(FPreFichier,'0000000'+
                               copy(HuitZero,1,(7-length(Trim(CodeClient))))+
                               Trim(CodeClient)+
                               ColleZeroDevant(NbCarnetLigne,3)+
                               ColleZeroDevant(NbTicketCarnet,2)+
                               ColleZeroDevant(NbTotalLigne,5)+
                               VFaciale+
                               PPatronale+
                               '                                '+
                               '                                '+
                              '2');
          NbCarnetW := NbCarnetW + 999;
        end;

        NbCarnetW := NbCarnetW - 999;
        Reste := NbT - (NbCarnetW*25);
        NbT := Reste ;
        if ((NbT div 25) * 25) < NbT then
          NbCarnetLigne := (NbT div 25)+1
        else
          NbCarnetLigne := (NbT div 25);
      end;

      if ((NbCarnetLigne * 25) > NbT) and (NbT > 25) then
      begin
        NbCarnetLigne := NbCarnetLigne - 1;
        NbTicketCarnet := 25;
        NbTotalLigne := NbCarnetLigne * NbTicketCarnet;
        Writeln(FPreFichier,'0000000'+
                            copy(HuitZero,1,(7-length(Trim(CodeClient))))+
                            Trim(CodeClient)+
                            ColleZeroDevant(NbCarnetLigne,3)+
                            ColleZeroDevant(NbTicketCarnet,2)+
                            ColleZeroDevant(NbTotalLigne,5)+
                            VFaciale+
                            PPatronale+
                            '                                '+
                            '                                '+
                            '2');
        NbCarnetLigne :=  1;
        NbTicketCarnet := NbT - NbTotalLigne;
        NbTotalLigne := NbTicketCarnet;
        Writeln(FPreFichier,'0000000'+
                            copy(HuitZero,1,(7-length(Trim(CodeClient))))+
                            Trim(CodeClient)+
                            ColleZeroDevant(NbCarnetLigne,3)+
                            ColleZeroDevant(NbTicketCarnet,2)+
                            ColleZeroDevant(NbTotalLigne,5)+
                            VFaciale+
                            PPatronale+
                            '                                '+
                            '                                '+
                            '2');

      end
      else
      begin
        NbTicketCarnet  := trunc(NbT / NbCarnetLigne);
        NbTotalLigne :=  NbTicketCarnet * NbCarnetLigne;
        Writeln(FPreFichier,'0000000'+
                            copy(HuitZero,1,(7-length(Trim(CodeClient))))+
                            Trim(CodeClient)+
                            ColleZeroDevant(NbCarnetLigne,3)+
                            ColleZeroDevant(NbTicketCarnet,2)+
                            ColleZeroDevant(NbTotalLigne,5)+
                            VFaciale+
                            PPatronale+
                            '                                '+
                            '                                '+
                            '2');
      end;
    end
// f PT1-2
// d PT2
    else
    if (Fournisseur = '003') then
    begin
    // ACCOR
      if (length(FormatFloat('0',Valeur(LaGrille.CellValues[1,i])*100)) = 5) then
        VFaciale := FormatFloat('0',Valeur(LaGrille.CellValues[1,i])*100)
      else

        VFaciale := copy(HuitZero,1,(5-length(FormatFloat('0',Valeur(LaGrille.CellValues[1,i])*100))))+
                 FormatFloat('0',Valeur(LaGrille.CellValues[1,i])*100);
        for j:= 1 to 5 do
        begin
          if (VFaciale[j] = '0')  then
            VFaciale[j] := ' '
          else
            break;
        end;
        if (VFaciale[5] = ' ') then
          VFaciale[5] := '0';


      if (length(FormatFloat('0',Valeur(LaGrille.CellValues[2,i])*100)) = 5) then
        PPatronale := FormatFloat('0',Valeur(LaGrille.CellValues[2,i])*100)
      else
        PPatronale := copy(HuitZero,1,(5-length(FormatFloat('0',Valeur(LaGrille.CellValues[2,i])*100))))+
                   FormatFloat('0',Valeur(LaGrille.CellValues[2,i])*100);
      for j:= 1 to 5 do
      begin
        if (PPatronale[j] = '0') then
          PPatronale[j] := ' '
        else
          break;
      end;
      if (PPatronale[5] = ' ') then
        PPatronale[5] := '0';

      ANbt := ColleZeroDevant(NbT,6);
      for j:= 1 to 6 do
      begin
        if (ANbt[j] = '0') then
          ANbt[j] := ' '
        else
          break;
      end;
      if (ANbt[6] = ' ') then
        ANbt[6] := '0';


      NiveauInterne1 := CodeDistribution;
      NiveauInterne2 := '      ';
      if (GetParamSocSecur ('SO_PGTKEDNPCARNET',FALSE) = TRUE) then
        EdNPCarnet := 'O'
      else
        EdNPCarnet := 'N';

      if(GetParamSocSecur ('SO_PGPERSOTICKET',FALSE) = TRUE) then
        EdNPTicket := 'O'
      else
        EdNPTicket := 'N';

      if (GetParamSocSecur ('SO_PGTKEDRSCARNET',FALSE) = TRUE) then
        EdRSCarnet := 'O'
      else
        EdRSCarnet := 'N';

      if (GetParamSocSecur ('SO_PGTKEDCPVCARNET',FALSE) = TRUE) then
        EdCPVCarnet := 'O'
      else
        EdCPVCarnet := 'N';

      CodePostal := GetParamSoc('SO_CODEPOSTAL');
      if (length(CodePostal) >= 5) then
        CodePostal :=  copy(CodePostal,1,5)
      else
        CodePostal := CodePostal + copy(blanc,1,(5-length(CodePostal)));

      Ville := GetParamSoc('SO_VILLE');
      if (length(Ville) >= 25) then
        Ville := copy(Ville,1,25)
      else
        Ville := Ville + copy(blanc,1,(25-length(Ville)));

      Writeln(FPreFichier,'50'+
                            ';'+
                            CodeClient+
                            ';'+
                            PointdeLivraison+
                            ';'+
                            NiveauInterne1+               // point de distribution
                            ';'+
                            NiveauInterne2+
                            ';'+
                            '-         '+
                            ';'+
                            '-                                       '+
                            ';'+
                            EdNPCarnet+
                            ';'+
                            EdNPTicket+
                            ';'+
                            trim(VFaciale)+
                            ';'+
                            trim(PPatronale)+
                            ';'+
                            IntToStr(NbT)+
                            ';'+
                            trim(Copy(RaisonSociale,1,34))+
                            ';'+
                            trim(CodePostal)+
                            ';'+
                            trim(Ville)+
                            ';'+
                            EdRSCarnet+
                            ';'+
                            EdCPVCarnet);
    end;
// f PT2
  end;
  {Enregitrement de fin de fichier}
  writeln(FPreFichier,'********'+
                      ColleZeroDevant(TotTicket,6)+
                      ColleZeroDevant(NbCarnet,6)+
                      copy(blanc,1,59)+
                      '3');

  CloseFile (FPreFichier);
end;

Initialization
  registerclasses ( [ TOF_PGTICKETSTOCK ] ) ;
end.
