{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 16/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGAFFPROFILPRES_MUL ()
Mots clefs ... : TOF;PGAFFPROFILPRES_MUL
*****************************************************************
PT1  17/08/2007  FLO  Ajout d'une suppression en masse des profils
PT2  27/08/2007  FLO  Date de recalcul = date de début de période
}
Unit UTofPG_MULAFFPROFILPRES ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     HDB,
     mul,
     fe_main,
{$else}
     eMul,
     uTob,
     maineagl,
{$ENDIF}
     Hqry,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF, htb97 ; 

Type
  TOF_PGAFFPROFILPRES_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    StDateArret : String;
    stFiltre : String;
    procedure OnClickSalarieSortie(Sender : TObject);
    procedure OnExitSalarie(Sender: TObject);
    procedure AssocListeSal(sender: Tobject);
    procedure OnChangeSALPROFIL(sender: Tobject);
    function GetXXWhere : String;
    procedure SupprimeAssoc (Sender: TObject); //PT1
    procedure OnDblClickListe(sender: Tobject);
    procedure Newassoc(sender: Tobject);
  end ;

Implementation
uses Entpaie, PGOutils2, P5Def, PGPresence;

procedure TOF_PGAFFPROFILPRES_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGAFFPROFILPRES_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGAFFPROFILPRES_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGAFFPROFILPRES_MUL.OnLoad ;
begin
  SetControlText('XX_WHERE',GetXXWhere);

     //PT2 - Début
     // Si visu des non associés, pas de suppression possible
     If GetControlText('SALPROFIL') = '002' Then
     Begin
          SetControlEnabled('BDELETE', False);
          // Changement de la liste associée
          If TFMul(Ecran).DBListe <> 'PGAFFPROFILPRESNO' Then
          Begin
               TFMul(Ecran).SetDBListe('PGAFFPROFILPRESNO');
               TFMul(Ecran).FListe.AllSelected := False;
               TFMul(Ecran).BChercheClick(Nil);
          End;
     End
     Else
     Begin
          SetControlEnabled('BDELETE', True);
          If TFMul(Ecran).DBListe <> 'PGAFFPROFILPRES' Then
          Begin
               TFMul(Ecran).SetDBListe('PGAFFPROFILPRES');
               TFMul(Ecran).FListe.AllSelected := False;
               TFMul(Ecran).BChercheClick(Nil);
          End;
     End;
     //PT2 - Fin
end;

procedure TOF_PGAFFPROFILPRES_MUL.OnArgument (S : String ) ;
var
  num : integer;
  Deb,Fin : TDateTime;
begin
  Inherited ;
     For Num := 1 to 4 do
          VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));

     (GetControl('CKSORTIE') as TCheckBox).OnClick := OnClickSalarieSortie;
     (GetControl('PSA_SALARIE') as ThEdit).OnExit:=OnExitSalarie;
     (GetControl('BOUVRIR') as TToolBarButton97).Onclick := AssocListeSal;
     OnClickSalarieSortie(Self);
     (GetControl('SALPROFIL') as THValComboBox).Value := '001'; // par défaut on affiche les salariés affectés
     (GetControl('SALPROFIL') as THValComboBox).OnChange := OnChangeSALPROFIL;
     (GetControl('BDELETE') as TToolBarButton97).Onclick := SupprimeAssoc; //PT1
     (GetControl('BINSERT') as TToolBarButton97).Onclick := Newassoc;

     {$IFNDEF EAGLCLIENT}
       (GetControl('FLISTE') as THDBGrid).OnDblClick := OnDblClickListe;
     {$ELSE}
       (GetControl('FLISTE') as THGrid).OnDblClick := OnDblClickListe;
     {$ENDIF}

     // Date de sortie des salariés : Début de période
     PresenceDonneMoisCalculActuel (Deb, Fin);
     SetControlText('DATEARRET', DateToStr(Deb));
end ;

procedure TOF_PGAFFPROFILPRES_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGAFFPROFILPRES_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGAFFPROFILPRES_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PGAFFPROFILPRES_MUL.OnClickSalarieSortie(Sender: TObject);
begin
  SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
  SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

procedure TOF_PGAFFPROFILPRES_MUL.OnExitSalarie(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGAFFPROFILPRES_MUL.AssocListeSal (sender : Tobject);
var
//  Q_Mul:THQuery ;
  Salarie, st : string;
  i: integer;
  Selectsal: string;
begin
  if (TFmul(Ecran).Fliste.NbSelected = 0) and (not TFmul(Ecran).Fliste.AllSelected) then
  begin
    PGIBOX('Aucun salarié n''est sélectionné.', Ecran.Caption);
    exit;
  end;
  if PgiAsk('Voulez-vous associer un nouveau profil de présence à tous les salariés sélectionnés ?', Ecran.caption) = mrYes then
  begin
    St := '';
    if (TFmul(Ecran).Fliste.AllSelected = TRUE) then
    begin // Si tout est sélectionné
      TFmul(Ecran).Q.First;
      while not TFmul(Ecran).Q.EOF do
      begin
      Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
      St := St + ' PSA_SALARIE="' + Salarie + '" OR';
      TFmul(Ecran).Q.Next;
      end;
      TFMul(Ecran).FListe.ClearSelected;
    end else begin // lecture de chaque salarié sélectionné
    { Composition du clause WHERE pour limiter le mul à ces salariés }
      for i := 0 to TFmul(Ecran).Fliste.NbSelected - 1 do
      begin
{$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
        TFMul(Ecran).Fliste.GotoLeBOOKMARK(i);
        Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
        St := St + Salarie + '|';
      end;
    end;
    Selectsal :=  St;
    AglLanceFiche('PAY', 'PGPROFILPRESSALGR', '', '', 'ACTION=CREATION-MULTISELECTION;'+Selectsal);
    TFMul(Ecran).FListe.ClearSelected;
    TFMul(Ecran).BCherche.Click;
  end;
end;


procedure TOF_PGAFFPROFILPRES_MUL.OnChangeSALPROFIL(sender: Tobject);
begin
  If GetControlText('SALPROFIL') = '002' then
  SetControlenabled('DATEAFFECT',false)
  else
  SetControlenabled('DATEAFFECT',true);
  
  SetControlText('XX_WHERE',GetXXWhere);
end;

function TOF_PGAFFPROFILPRES_MUL.GetXXWhere: String;
var
  DateArret, Dateaffect : TDateTime;
  stdateaffect : string;
begin
  Inherited ;
  StDateArret:='';
  stdateaffect := '';

  if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then
  Begin
    DateArret:=StrtoDate(GetControlText('DATEARRET'));
    StDateArret:=' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) '+
    ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
  End;

  if Getcontroltext('SALPROFIL') = '001' then
  begin
    Dateaffect := Strtodate(GetcontrolText('DATEAFFECT'));
    stdateaffect := ' AND PPZ_DATEVALIDITE in (SELECT MAX(PPZ_DATEVALIDITE) FROM PROFILPRESSALARIE e' +
    ' WHERE e.PPZ_DATEVALIDITE <="'+UsDateTime(DateAffect)+'" AND PSA_SALARIE = e.PPZ_SALARIE)';
  end;

   result := StDateArret + stDateAffect + stFiltre;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 17/08/2007
Modifié le ... :   /  /    
Description .. : Suppression en masse d'associations de profils
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGAFFPROFILPRES_MUL.SupprimeAssoc(Sender: TObject);
Var Salarie, St : String;
    i : Integer;
    DD,DF : TDateTime;
begin
     // Elements sélectionnés
     If (TFmul(Ecran).Fliste.NbSelected = 0) And (Not TFmul(Ecran).Fliste.AllSelected) Then
     Begin
          PGIBox(TraduireMemoire('Aucun salarié n''est sélectionné.'), TraduireMemoire('Suppression des profils'));
          Exit;
     End;

     If PgiAsk(TraduireMemoire('Voulez-vous supprimer les profils de présence associés aux salariés sélectionnés ?'), Ecran.caption) = mrYes Then
     Begin
          St := '';

          // Si tout est sélectionné
          If (TFmul(Ecran).Fliste.AllSelected = True) Then
          Begin
               TFmul(Ecran).Q.First;
               While Not TFmul(Ecran).Q.EOF Do
               Begin
                    If St <> '' Then St := St + ',';
                    Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').AsString;
                         St := St + '"' + Salarie + '"';
                    TFmul(Ecran).Q.Next;
               End;
          End
          // Une partie des salariés a été sélectionnée
          Else
          Begin
               For i := 0 to TFmul(Ecran).Fliste.NbSelected - 1 Do
               Begin
                    If St <> '' Then St := St + ',';
                    {$IFDEF EAGLCLIENT}
                         TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
                    {$ENDIF}
                    TFMul(Ecran).Fliste.GotoLeBOOKMARK(i);
                    Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').AsString;
                         St := St + '"' + Salarie + '"';
               End;
          End;

          If St <> '' Then
          Begin
               ExecuteSQL('DELETE FROM PROFILPRESSALARIE WHERE PPZ_SALARIE IN ('+St+')');

               // Mise en place des tops de recalcul pour chaque salarié
               If (LastError = 0) Then
               Begin
                    PresenceDonneMoisCalculActuel (DD,DF); //PT2
                    Salarie := READTOKENPipe(St, ',');
                    While Salarie <> '' Do
                    Begin
                         // Suppression des guillemets
                         Salarie := Copy (Salarie, 2, Length(Salarie)-2);
                         CompteursARecalculer(DD, Salarie); //PT2
                    End;
               End;
          End;

          TFMul(Ecran).FListe.ClearSelected;
          TFMul(Ecran).BCherche.Click;
     End;
end;

procedure TOF_PGAFFPROFILPRES_MUL.OnDblClickListe(sender: Tobject);
var Q_Mul:THQuery ;
begin
    Q_Mul:=THQuery(Ecran.FindComponent('Q'));

    
   {$IFDEF EAGLCLIENT}
    TFMul (Ecran).Q.TQ.Seek (TFMul (Ecran).FListe.Row-1) ;
   {$ENDIF}

    If GetControlText('SALPROFIL') = '001' Then
      AglLanceFiche('PAY', 'PGPROFILPRESSAL', '', Q_MUL.FindField('PSA_SALARIE').AsString + ';' + Q_MUL.FindField('PPZ_DATEVALIDITE').AsString, 'ACTION=MODIFICATION;'+Q_MUL.FindField('PSA_SALARIE').AsString + ';' + Q_MUL.FindField('PPZ_DATEVALIDITE').AsString)
    else
      AglLancefiche('PAY', 'PGPROFILPRESSAL','','','ACTION=CREATION;'+ Q_MUL.FindField('PSA_SALARIE').AsString +';');
    TFMul(Ecran).BCherche.Click;
end;

procedure TOF_PGAFFPROFILPRES_MUL.Newassoc(sender: Tobject);
begin
    AglLancefiche('PAY', 'PGPROFILPRESSAL','','','ACTION=CREATION');
end;

Initialization
  registerclasses ( [ TOF_PGAFFPROFILPRES_MUL ] ) ; 
end.
