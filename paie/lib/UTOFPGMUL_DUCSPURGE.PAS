{***********UNITE*************************************************
Auteur  ...... : GG
Créé le ...... : 29/12/2006
Modifié le ... :   /  /
Description .. : Multicritère de sélection des DUCS en vue de purge des anciennes déclarations
Mots clefs ... : TOF;PGMUL_DUCSPURGE
*****************************************************************}
{
PT1 02/02/2007 : FC Suivant un paramètre société, contrer l'habilitation qui est faite automatique
                    en effaçant le contenu du champ XX_WHERE2
PT2 25/04/2007 : MF Complément sur message de suppression d'une déclaration de
                    l'exercice
PT3 19/07/2007 : MF FQ 14533 Inscription de la purge des Ducs au jnal des événements
}
Unit UTofPGMul_DUCSPurge ;

Interface

uses

  {$IFDEF VER150}
  Variants,
  {$ENDIF}
{$IFDEF EAGLCLIENT}
  UtileAGL,
  eMul,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  HDB,
  Mul,
{$ENDIF}

  Classes, sysutils, HTB97, hmsgbox, HCtrls, Hqry, UTOF, ParamDat,
  Hent1, DateUtils, Controls,  hstatus, ed_tools,
  ParamSoc,//PT1
  P5Def;      // PT3 pour le jnal des événements

Type
  TOF_PGMUL_DUCSPURGE = Class (TOF)
  public
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
    WW: THEdit;
    DateDeb, DateFin: THEdit;
    Q_Mul: THQuery;
{$IFNDEF EAGLCLIENT}
    Liste : THDBGrid;
{$ELSE}
    Liste : THGrid;
{$ENDIF}
    procedure ActiveWhere;
    procedure DateElipsisclick(Sender: TObject);
    procedure DeleteClick(Sender: TObject);
    procedure ExecutePurgeUn(Trace : TstringList);
    function ExerciceEnCoursInclue(Deb, Fin : TDateTime): Boolean;
  end ;

Implementation

uses PgOutils, PgOutils2, DB;

function TOF_PGMUL_DUCSPURGE.ExerciceEnCoursInclue(Deb, Fin : TDateTime) : Boolean;
var
  ExerPerEncours, MoisE, AnneeE : string;
  DebExer, FinExer : TDateTime;
begin
  result := False;
  if RendExerSocialEnCours(MoisE, AnneeE, ExerPerEncours, DebExer, FinExer) = True then
  begin
    // Si l'exercice en cours recoupe les dates sélectionnées
    if (   ((DebExer >= Deb) and (DebExer <= Fin))
        or ((FinExer >= Deb) and (FinExer <= Fin))
        or ((DebExer <= Deb) and (FinExer >= Fin))) then
        begin
          result := True;
        end;
  end;
end;

procedure TOF_PGMUL_DUCSPURGE.OnLoad ;
begin
  Inherited ;
  ActiveWhere;
end ;

procedure TOF_PGMUL_DUCSPURGE.OnArgument (S : String ) ;
var
  ExerPerEncours, MoisE, AnneeE : string;
  DebExer, FinExer : TDateTime;
  Delete : TToolbarButton97;
begin
  inherited;
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  WW := THEdit(GetControl('XX_WHERE'));
  DateDeb := ThEdit(getcontrol('XX_VARIABLED'));
  DateFin := ThEdit(getcontrol('XX_VARIABLED_'));
  if (DateDeb <> nil) and (DateFin <> nil) then
  begin
    DateDeb.OnElipsisClick := DateElipsisclick;
    DateFin.OnElipsisClick := DateElipsisclick;
  end;
  // Date par défaut
  // La date proposée est L'exercice en cours moins 1 an.
  if RendExerSocialEnCours(MoisE, AnneeE, ExerPerEncours, DebExer, FinExer) = True then
  begin
     if Assigned(DateDeb) then DateDeb.text := DateToStr(IncYear(DebExer,-1));
     if Assigned(DateFin) then DateFin.text := DateToStr(IncYear(FinExer,-1));
  end;
  Delete := TToolbarButton97(GetControl('BDelete'));
  if Assigned(Delete) then
  begin
    Delete.OnClick := DeleteClick;
  end;
end ;

procedure TOF_PGMUL_DUCSPURGE.ActiveWhere;
var
  St: string;
  Where2:THEdit;//PT1
begin
  WW.Text := '';
  st := '';
  St := '(PDU_DATEDEBUT >="' + UsDateTime(StrToDate(DateDeb.text)) + '") AND (PDU_DATEDEBUT <="' + UsDateTime(StrToDate(DateFin.text)) +
    '") AND (PDU_DATEFIN >="' + UsDateTime(StrToDate(DateDeb.text)) + '") AND (PDU_DATEFIN <= "' + UsDateTime(StrToDate(DateFin.text)) + '")';
  if St <> '' then WW.Text := st;
  if Q_Mul <> nil then
  begin
    TFMul(Ecran).SetDBListe('PGDUCSENTETE');
  end;

  //DEB PT1
  if GetParamSocSecur('SO_PGDRTVISUETAB',True) then
  begin
    Where2 := THEdit(GetControl('XX_WHERE2'));
    if Where2 <> nil then SetControlText('XX_WHERE2', '');
  end;
  //FIN PT1
end;

procedure TOF_PGMUL_DUCSPURGE.DateElipsisclick(Sender: TObject);
var key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGMUL_DUCSPURGE.DeleteClick(Sender: TObject);
var
  reponse, i : integer;
  Trace, TraceE :  TStringList; // PT3
begin
{$IFNDEF EAGLCLIENT}
  Liste:= THDBGrid(GetControl('FListe'));
{$ELSE}
  Liste:= THGrid(GetControl('FListe'));
{$ENDIF}
  reponse:= PGIAsk ('Attention, vous allez supprimer#13#10'+
                    'une partie de l''archivage de vos déclarations.#13#10'+
                    'Voulez-vous continuer ?',
                    TFMul (Ecran).Caption);
  if reponse = mrYes then
  begin
    if Liste <> nil then
    begin
      if ((Liste.NbSelected=0) and (not Liste.AllSelected)) then
      begin
        MessageAlerte('Aucun élément sélectionné');
        exit;
      end;
// d PT3
      Trace := TStringList.Create;
      TraceE := TStringList.Create;
// f PT3
      if (Liste.AllSelected = TRUE) then
      begin
        Trace.add('Traitement de purge des DUCS'); //PT3
        InitMoveProgressForm (NIL,'Suppression en cours', 'Veuillez patienter SVP ...',
                              Liste.NbSelected,FALSE,TRUE);
        InitMove(TFmul(Ecran).Q.RecordCount, '');
        TFmul(Ecran).Q.First;
        while not TFmul(Ecran).Q.EOF do
        begin
          ExecutePurgeUn(Trace); //PT3
          TFmul(Ecran).Q.Next;
        end;
        Liste.AllSelected := False;
        FiniMove;
        FiniMoveProgressForm;
        TFMul(Ecran).bSelectAll.Down := False;
        //On rafraichi l'affichage de la liste de sélection pour enlever les éléments supprimés
        TFMul(Ecran).BChercheClick(Sender);
      end
      else // Tout n'est pas sélectionné
      begin
        Trace.add('Traitement de purge des DUCS'); //PT3
        InitMoveProgressForm (NIL,'Suppression en cours', 'Veuillez patienter SVP ...',
                              Liste.NbSelected,FALSE,TRUE);
        InitMove(Liste.NbSelected, '');
        try
          for i:=0 to Liste.NbSelected-1 do
          begin
            Liste.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
            TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
{$ENDIF}
            ExecutePurgeUn(Trace); //PT3
          end;
        finally
          Liste.ClearSelected;
          FiniMove;
          FiniMoveProgressForm;
          //On rafraichi l'affichage de la liste de sélection pour enlever les éléments supprimés
          TFMul(Ecran).BChercheClick(Sender);
        end;
      end;
// d PT3
      Trace.add('Fin du traitement de purge des DUCS');
      CreeJnalEvt ('001','023','Ok',nil,nil,Trace, TraceE);
      Trace.free;
      TraceE.free;
// f PT3
    end;
  end;
end;

procedure TOF_PGMUL_DUCSPURGE.ExecutePurgeUn(Trace :Tstringlist);
var
  PDU_ETABLISSEMENT, PDU_ORGANISME : String;
  PDU_DATEDEBUT, PDU_DATEFIN : TDateTime;
  PDU_NUM : Integer;
  bExerciceEnCoursInclue : Boolean;
begin
   PDU_ETABLISSEMENT:= TFmul (Ecran).Q.FindField('PDU_ETABLISSEMENT').asstring;
   PDU_ORGANISME:= TFmul (Ecran).Q.FindField('PDU_ORGANISME').asstring;
   PDU_DATEDEBUT:= TFmul (Ecran).Q.FindField('PDU_DATEDEBUT').AsDateTime;
   PDU_DATEFIN:= TFmul (Ecran).Q.FindField('PDU_DATEFIN').AsDateTime;
   PDU_NUM:= TFmul (Ecran).Q.FindField('PDU_NUM').AsInteger;
   bExerciceEnCoursInclue :=  exerciceEnCoursInclue(PDU_DATEDEBUT,PDU_DATEFIN);
   if ((bExerciceEnCoursInclue and (PGIAsk (
                                    'Attention, la déclaration sélectionnée concerne l''exercice en cours#13#10'+
// d PT2
                                    'Etablissement : '+
                                    TFmul (Ecran).Q.FindField('PDU_ETABLISSEMENT').asstring+
                                    ' Organisme : '+
                                    TFmul (Ecran).Q.FindField('PDU_ORGANISME').asstring+
                                    ' du ' + DateToStr(TFmul (Ecran).Q.FindField('PDU_DATEDEBUT').AsDateTime)+
                                    ' au ' + DateToStr(TFmul (Ecran).Q.FindField('PDU_DATEFIN').AsDateTime)+
                                    '#13#10'+
// f PT2
                                    'Voulez-vous la supprimer ?', TFMul (Ecran).Caption)= mrYes) )
        or not bExerciceEnCoursInclue ) then
      begin
      if ((PDU_ETABLISSEMENT <> '') and (PDU_ORGANISME <> '')) then
         begin
         //On supprime le contenu avant de supprimer l'entête
         ExecuteSQL('DELETE FROM DUCSDETAIL WHERE PDD_ETABLISSEMENT = "'+PDU_ETABLISSEMENT+'"'
                                           +' AND PDD_ORGANISME = "'+PDU_ORGANISME+'"'
                                           +' AND PDD_DATEDEBUT = "'+UsDateTime (PDU_DATEDEBUT)+'"'
                                           +' AND PDD_DATEFIN = "'+UsDateTime (PDU_DATEFIN)+'"'
                                           +' AND PDD_NUM = '+inttostr(PDU_NUM));
         ExecuteSQL('DELETE FROM DUCSENTETE WHERE PDU_ETABLISSEMENT = "'+PDU_ETABLISSEMENT+'"'
                                           +' AND PDU_ORGANISME = "'+PDU_ORGANISME+'"'
                                           +' AND PDU_DATEDEBUT = "'+UsDateTime (PDU_DATEDEBUT)+'"'
                                           +' AND PDU_DATEFIN = "'+UsDateTime (PDU_DATEFIN)+'"'
                                           +' AND PDU_NUM = '+inttostr(PDU_NUM));
// d PT3
         Trace.add('suppression Etab '+ TFmul (Ecran).Q.FindField('PDU_ETABLISSEMENT').asstring+
                   ' organisme '+ TFmul (Ecran).Q.FindField('PDU_ORGANISME').asstring+
                   ' période du '+DateToStr(TFmul (Ecran).Q.FindField('PDU_DATEDEBUT').AsDateTime)+
                   ' au '+DateToStr(TFmul (Ecran).Q.FindField('PDU_DATEFIN').AsDateTime));
// f PT3
         end;
      MoveCur(False);
      MoveCurProgressForm;
      end;
end;

Initialization
  registerclasses ( [ TOF_PGMUL_DUCSPURGE ] ) ;
end.
