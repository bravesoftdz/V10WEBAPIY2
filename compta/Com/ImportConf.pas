unit ImportConf;

interface
uses
  Windows, SysUtils, Classes, Controls, Forms,
  hmsgbox,
{$IFNDEF EAGLCLIENT}
  DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  Hctrls, HEnt1, Ent1, HTB97, stdCtrls,
  UTOF,Vierge, UtilTrans,Soldecpt, UTob, UtilGed;

type
  TEvtSuptot = class
    Dateecr1,Dateecr2: TDateTime;
    Jr,Exo           : string;
    Origine          : string;
    SuppreANo        : Boolean;
    procedure DegageEcrituretotal;
    procedure DegageJournal;
    procedure DegageExercice;
    procedure DegageDocumentGed(Where : string);
  end;

type
  TOF_ImportConf = Class (TOF)
  procedure OnLoad ; override ;
  procedure OnArgument(stArgument: String); override ;
  procedure OnClose ; override ;
  private
  Nature    : string;
  Origine   : string;
  AppelEchange     : Boolean;
  Result    : string;
  procedure bClickValide(Sender: TObject);
  procedure PChange(Sender: TObject);
  Procedure AnoOnClick(Sender :TObject) ;
  Procedure ExerciceOnChange(Sender :TObject) ;
  procedure OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
end;

implementation


uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  Paramsoc;

procedure TOF_ImportConf.OnArgument(stArgument: String);
var
St        : string;
indice    : integer;
Q         : TQuery;
RAZ       : Boolean;
x,x2      : string;
Codexo    : string;
Q1        : TQuery;
begin
        indice := 1;  RAZ := FALSE;
        St := ReadTokenSt(stArgument);
        if Pos ('Appel : ECHANGES PGI', stArgument ) <> 0 then
           AppelEchange := TRUE
        else
           AppelEchange := FALSE;

        while St <> '' do
        begin
            if indice = 3 then
            begin
                 x2 := St;
                 x := ReadTokenPipe (x2,'\');
                 while x <> '' do
                 begin
                      x := ReadTokenPipe (x2,'\');
                      if x <> '' then
                       St := 'Fichier : ' + x;
                 end;
                 Codexo := ReadTokenSt(stArgument);
            end;
            if indice = 1 then Origine := Copy (St,22, Length(St)-21);
             SetControlText ('EDIT'+IntToStr(indice), St);
             St := ReadTokenSt(stArgument);
             if copy (St, 1, 10) = ' Nature : ' then
                Nature := copy (St, 11, 3);
             if copy (St, 1, 3) = 'RAZ' then RAZ := TRUE;
             inc (indice);
        end;
     TToolbarButton97(GetControl('BValider')).Onclick := bClickValide;
    THMultiValComboBox(GetControl('JOURNAUX')).Items.clear;
    THMultiValComboBox(GetControl('JOURNAUX')).Values.clear;
    if (Origine = 'S1')  and (GetparamsocSecur ('SO_CPMODESYNCHRO', TRUE) = FALSE) then TCheckBox(GetControl('SUPPANOUVEAU')).Checked := TRUE;

    if (Origine <> 'S1') or (TCheckBox(GetControl('SUPPANOUVEAU')).Checked = FALSE) then
    Q := OpenSQL('select J_JOURNAL,J_LIBELLE,J_NATUREJAL FROM JOURNAL '+
    'Where J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"CLO" ORDER BY J_JOURNAL' ,true)
    else
    Q := OpenSQL('select J_JOURNAL,J_LIBELLE,J_NATUREJAL FROM JOURNAL ORDER BY J_JOURNAL' ,true);

    while not Q.EOF do
    begin
         THMultiValComboBox(GetControl('JOURNAUX')).Items.Add(Q.FindField('J_LIBELLE').AsString);
         THMultiValComboBox(GetControl('JOURNAUX')).Values.add(Q.FindField('J_JOURNAL').AsString);
         Q.Next;
    end;
    Ferme(Q);

     if (Nature = 'DOS') then // or (not (TCheckBox(GetControl('SUPPEXE')).Checked)) then
     begin
          TCheckBox(GetControl('SUPPRTOTAL')).Checked := TRUE;
          SetControlEnabled ('SUPPRTOTAL', FALSE);
          SetControlEnabled ('FE__GROUPBOX1', FALSE);
          SetControlEnabled ('SUPPEXE',FALSE);
          TLabel(GetControl('TEXERC')).Enabled       := false;
          TLabel(GetControl('TJOURNAL')).Enabled     := false;
          TLabel(GetControl('TDATEECR')).Enabled     := false;
          TLabel(GetControl('TDATEECR1')).Enabled    := false;
     end
     else
     begin
          TCheckBox(GetControl('SUPPRTOTAL')).Checked := FALSE;
          SetControlText ('DATEECR', stDate1900);
          SetControlText ('DATEECR1', stDate2099);
          PChange(Ecran);
     end;
    if RAZ = TRUE then
    begin
          TCheckBox(GetControl('SUPPRTOTAL')).Checked := TRUE;
          SetControlEnabled ('SUPPRTOTAL', FALSE);
          SetControlEnabled ('SUPPEXE', FALSE);
          SetControlEnabled ('FE__GROUPBOX1', FALSE);
          TCheckBox(GetControl('SUPPEXE')).Checked := FALSE;
    end;

     TCheckBox(GetControl('SUPPRTOTAL')).OnClick := PChange;
     TCheckBox(GetControl('SUPPEXE')).OnClick := PChange;
     THValComboBox(GetControl('EXERCICE')).OnChange := ExerciceOnChange;
     TCheckBox(GetControl('SUPPANOUVEAU')).OnClick  := AnoOnClick;
     Ecran.OnKeyDown := OnKeyDownEcran;
  if (Nature = 'BAL') then
  begin
    SetControlChecked('SUPPRTOTAL',False);
    SetControlEnabled('SUPPRTOTAL',False);
  end;
  // Si Nature = Journal , on ne peut pas faire de RAZ dossier
     // ajout me 9-12-2002 SISCO EXO
  if (Nature = 'SYN') or (Nature = 'JRL') or (((Nature = 'BAL') or (Nature = 'EXE')) and (Origine= 'SI')) then
  begin
    SetControlChecked('SUPPRTOTAL',False);
    SetControlEnabled('SUPPRTOTAL',False);
    if (Nature = 'SYN') then
    begin
         SetControlEnabled ('SUPPEXE', FALSE);
         SetControlEnabled ('FE__GROUPBOX1', FALSE);
    end;
    if (Nature = 'EXE') then
    begin
          TLabel(GetControl('TJOURNAL')).Enabled     := false;
          TLabel(GetControl('TDATEECR')).Enabled     := false;
          TLabel(GetControl('TDATEECR1')).Enabled    := false;
          TCheckBox(GetControl('SUPPEXE')).Checked   := TRUE;
          THVALComboBox(GetControl('EXERCICE')).Items.clear;
          THVALComboBox(GetControl('EXERCICE')).Values.clear;
          Q1 := OpenSql ('SELECT EX_LIBELLE,EX_EXERCICE from EXERCICE WHERE EX_EXERCICE="'+Codexo+'"',FALSE);
          if not Q1.EOF  then
          begin
                   THVALComboBox(GetControl('EXERCICE')).Items.Add(Q1.FindField('EX_LIBELLE').asstring);
                   THVALComboBox(GetControl('EXERCICE')).Values.add(Q1.FindField('EX_EXERCICE').asstring);
                   THVALComboBox(GetControl('EXERCICE')).ItemIndex := 0;
                   THVALComboBox(GetControl('EXERCICE')).Enabled :=  FALSE;
          end;
          ferme (Q1);
    end;
  end;
  if (GetParamSocSecur ('SO_CPLIENGAMME','') = 'S1') and (Origine = 'S1') then
            SetControlVisible ('SUPPANOUVEAU', TRUE)
  else
            SetControlVisible ('SUPPANOUVEAU', FALSE);
  SetControlEnabled ('SUPPANOUVEAU', FALSE);

  // Mise à jour du récapitulatif
  if Origine = 'SI' then Origine := 'Sisco II';
  SetControlText ('EDIT1','Origine du fichier : '+ Origine);
  if Nature = 'JRL' then Nature := 'Journal' else
  if Nature = 'BAL' then Nature := 'Balance' else
  if Nature = 'DOS' then Nature := 'Dossier' else
  if Nature = 'SYN' then Nature := 'Synchronisation Business Line';
  SetControlText ('EDIT1','Origine du fichier : '+ Origine);
  SetControlText ('EDIT2','Nature du fichier : '+ Nature);

end;

procedure TOF_ImportConf.OnLoad ;
begin
    Result := '-';
end;

procedure TOF_ImportConf.OnClose ;
begin
     TFVierge (Ecran).retour := Result;
end;

procedure TEvtSuptot.DegageEcrituretotal;
begin
        DegageDocumentGed('');
        ExecuteSQL('DELETE FROM EXERCICE');
        ExecuteSQL('DELETE FROM ECRITURE');
        ExecuteSQL('DELETE FROM ANALYTIQ');
        ExecuteSQL('DELETE FROM ECRCOMPL');
        ExecuteSQL('DELETE FROM EEXBQ');
        ExecuteSQL('DELETE FROM EEXBQLIG');
        ExecuteSQl('DELETE FROM CROISEAXE');
        // bap
        ExecuteSQL('DELETE FROM CPTACHEBAP');
        ExecuteSQL('DELETE FROM CPBONSAPAYER');
        MajTotTousComptes(False, '');
end;


procedure TEvtSuptot.DegageDocumentGed(Where : string);
var
Q1            : TQuery;
ii            : integer;
TPieceCon     : TOB;
Wheres        : string;
begin

(*   Fiche 17025
          if Where = '' then Wheres := ' Where e_docid <> 0 '
             else  Wheres := Where + ' and e_docid <> 0 ';

             Q1 := Opensql ('select e_docid from ecriture '+Wheres+' group by e_docid', TRUE);
             TPieceCon := TOB.Create('', nil, -1);
             TPieceCon.LoadDetailDB('Ecriture', '', '', Q1, TRUE, FALSE);
             Ferme(Q1);
             for  ii := 0  to  TPieceCon.detail.count-1 do
                 SupprimeDocumentGed (TPieceCon.detail[ii].Getvalue ('E_DOCID'));
             TPieceCon.free;
*)
             if Where = '' then Wheres := ' Where EC_DOCGUID <> "" '
             else  Wheres := Where + ' and EC_DOCGUID <> "" ';

             Q1 := Opensql ('select EC_DOCGUID from ECRCOMPL '+Wheres+' group by EC_DOCGUID', TRUE);
             TPieceCon := TOB.Create('', nil, -1);
             TPieceCon.LoadDetailDB('ECRCOMPL', '', '', Q1, TRUE, FALSE);
             Ferme(Q1);
             for  ii := 0  to  TPieceCon.detail.count-1 do
                 SupprimeDocumentGed (TPieceCon.detail[ii].Getvalue ('EC_DOCGUID'));
             TPieceCon.free;

end;

procedure TEvtSuptot.DegageJournal;
var
Where, Wheres : string;
begin
        if (Origine <> 'S1') or (SuppreANo = FALSE)  then
        Where := 'Where E_ECRANOUVEAU<>"OAN" AND E_ECRANOUVEAU<>"H"'
        else
        Where := 'Where ';

        if (Exo <> '') then
                Where := Where + ' AND E_EXERCICE="'+Exo + '"' ;
        if (Jr <> '') then
                Where := Where + ' AND (' + RendCommandeJournal('E',Jr)+')';
        Where := Where + ' AND (E_DATECOMPTABLE>="'+USDateTime(Dateecr1)+'"' +
                              ' AND E_DATECOMPTABLE<="'+USDateTime(Dateecr2)+'")';
        Wheres := Where + ' AND (E_ETATLETTRAGE = "TL"' +
        ' OR  E_ETATLETTRAGE = "PL" OR E_REFPOINTAGE<>"")';

        if ExisteSQl ('SELECT * from ecriture '+ Wheres ) then
        begin
             PGIInfo (' Il existe des écritures lettrées et/ou pointées', 'Impossible de supprimer les écritures');
             V_PGI.IoError := oeUnknown;
        end
        else
            ExecuteSQL('DELETE FROM ECRITURE '+ Where);

        Where := '';
        if (Exo <> '') then
                Where := ' Where Y_EXERCICE="'+Exo + '"' ;
        if (Jr <> '') then
        begin
                if Where = '' then   Where := ' Where (' + RendCommandeJournal('Y',Jr)+')'
                else
                Where := Where + ' AND (' + RendCommandeJournal('Y',Jr)+')';
        end;
        if Where = '' then
           Where :=  ' Where (Y_DATECOMPTABLE>="'+USDateTime(Dateecr1)+'"' +
                              ' AND Y_DATECOMPTABLE<="'+USDateTime(Dateecr2)+'")'
        else
           Where := Where + ' AND (Y_DATECOMPTABLE>="'+USDateTime(Dateecr1)+'"' +
                              ' AND Y_DATECOMPTABLE<="'+USDateTime(Dateecr2)+'")';
        ExecuteSQL('DELETE FROM ANALYTIQ '+ Where);

        Where := ''; // suppression des écritures complémentaires
        if (Exo <> '') then
                Where := ' Where  EC_EXERCICE="'+Exo + '"' ;
        if (Jr <> '') then
        begin
                if Where = '' then
                   Where := ' Where (' + RendCommandeJournal('EC',Jr)+')'
                else
                   Where := Where + ' AND (' + RendCommandeJournal('EC',Jr)+')';
        end;
        if Where = '' then
           Where := ' Where (EC_DATECOMPTABLE>="'+USDateTime(Dateecr1)+'"' +
                              ' AND EC_DATECOMPTABLE<="'+USDateTime(Dateecr2)+'")'
        else
           Where := Where + ' AND (EC_DATECOMPTABLE>="'+USDateTime(Dateecr1)+'"' +
                              ' AND EC_DATECOMPTABLE<="'+USDateTime(Dateecr2)+'")';
        DegageDocumentGed (Where);   // Fiche 17025
        ExecuteSQL('DELETE FROM ECRCOMPL '+ Where);

        Where := ''; // suppression bon à payer
        if (Exo <> '') then
                Where := ' Where  BAP_EXERCICE="'+Exo + '"' ;
        if (Jr <> '') then
        begin
                if Where = '' then
                   Where := ' Where (' + RendCommandeJournal('BAP',Jr)+')'
                else
                   Where := Where + ' AND (' + RendCommandeJournal('BAP',Jr)+')';
        end;
        if Where = '' then
           Where := ' Where  (BAP_DATECOMPTABLE>="'+USDateTime(Dateecr1)+'"' +
                              ' AND BAP_DATECOMPTABLE<="'+USDateTime(Dateecr2)+'")'
        else
           Where := Where + ' AND (BAP_DATECOMPTABLE>="'+USDateTime(Dateecr1)+'"' +
                              ' AND BAP_DATECOMPTABLE<="'+USDateTime(Dateecr2)+'")';
        ExecuteSQL('DELETE FROM CPBONSAPAYER '+ Where);

        MajTotTousComptes(False, '');

end;

     // ajout me 9-12-2002 SISCO EXO
procedure TEvtSuptot.DegageExercice;
var
Wheres : string;
begin
        Wheres := 'Where E_EXERCICE="'+Exo + '"';
        ExecuteSQL('DELETE FROM ECRITURE '+ Wheres);

        Wheres := ' Where Y_EXERCICE="'+Exo + '"' ;
        ExecuteSQL('DELETE FROM ANALYTIQ '+ Wheres);
        // bap

        Wheres := ' Where BAP_EXERCICE="'+Exo + '"' ;
        ExecuteSQL('DELETE FROM CPBONSAPAYER');


        MajTotTousComptes(False, '');

end;

procedure TOF_ImportConf.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F10    :  begin bClickValide(Sender); Ecran.ModalResult := 1; end;
  end;
end;


procedure TOF_ImportConf.bClickValide(Sender: TObject);
var
  EvtSup        : TEvtSuptot;
begin

     if TCheckBox(GetControl('SUPPRTOTAL')).Checked then
     begin
          If HShowMessage('0;Remise à zéro des écritures; Confirmez-vous la suppression des écritures du dossier ?;Q;YNC;N;C;','','')<>mrYes Then
          begin Result := '-'; exit; end;

      EvtSup := TEvtSuptot.Create;
      // GP : pb transaction imbriquées avec begintrans. fait pour client secto HERMES : ca urgeait les baguouinsse...
      (*
      if Transactions(EvtSup.DegageEcrituretotal, 1) <> oeOK then
      begin
          MessageAlerte('Suppression Totale impossible'); Result := '-'; EvtSup.Free; exit;
      end;
      *)
      EvtSup.DegageEcrituretotal ;
      EvtSup.Free;
     end;
     if TCheckBox(GetControl('SUPPEXE')).Checked then
     begin
          If HShowMessage('0;Remise à zéro Exercice; Confirmez-vous la suppression de l''exercice ?;Q;YNC;N;C;','','')<>mrYes Then
          begin Result := '-'; exit; end;

          EvtSup := TEvtSuptot.Create;
          EvtSup.Jr := ''; EvtSup.Exo := '';
          EVtSup.Origine := Origine;
          EvtSup.Dateecr1         := StrToDate(GetControlText('DATEECR'));
          EvtSup.Dateecr2         := StrToDate(GetControlText('DATEECR1'));
//GP          if (THMultiVALComboBox(GetControl('Journaux')).Text <> '<<Tous>>')
          if (THMultiVALComboBox(GetControl('Journaux')).Text <> Traduirememoire('<<Tous>>'))
          and (THMultiVALComboBox(GetControl('Journaux')).Text <> '') then
           EvtSup.Jr            := THMultiVALComboBox(GetControl('Journaux')).Text;
//GP          if (THVALComboBox(GetControl('EXERCICE')).Text <> '<<Tous>>')
          if (THVALComboBox(GetControl('EXERCICE')).Text <> Traduirememoire('<<Tous>>'))
           and (THVALComboBox(GetControl('EXERCICE')).Text <> '') then
           EvtSup.Exo           := THVALComboBox(GetControl('EXERCICE')).value;
          EvtSup.SuppreANo      := TCheckBox(GetControl('SUPPANOUVEAU')).Checked;
          if (EvtSup.Jr = '') and  (EvtSup.Exo = '') then
          begin
               // MessageAlerte('Veuillez effectuer la RAZ des écritures du dossier');  Result := '-'; EvtSup.Free; exit;
          end
          else
          begin
               if (Nature = 'EXE') and (Origine = 'Sisco II') then
               begin
                  if Transactions(EvtSup.DegageExercice, 0) <> oeOK then
                  begin
                       MessageAlerte('Suppression exercice impossible'); Result := '-'; EvtSup.Free; exit;
                  end;
               end
               else
               begin
                  if Transactions(EvtSup.DegageJournal, 0) <> oeOK then
                  begin
                      MessageAlerte('Suppression exercice impossible'); Result := '-'; EvtSup.Free; exit;
                  end;
               end;
          end;
          EvtSup.Free;
     end;
     Result := 'X';
     if AppelEchange and (not TCheckBox(GetControl('SUPPEXE')).Checked) then
        Result := '?';
end;

procedure TOF_ImportConf.PChange(Sender: TObject);
begin
    SetControlEnabled ('SUPPEXE', not TCheckBox(GetControl('SUPPRTOTAL')).Checked);
    SetControlEnabled ('FE__GROUPBOX1', not TCheckBox(GetControl('SUPPRTOTAL')).Checked);

    SetControlEnabled ('SUPPRTOTAL',not TCheckBox(GetControl('SUPPEXE')).Checked);
    SetControlEnabled ('FE__GROUPBOX1', TCheckBox(GetControl('SUPPEXE')).Checked);
    TLabel(GetControl('TEXERC')).Enabled       := TCheckBox(GetControl('SUPPEXE')).Checked;
    TLabel(GetControl('TJOURNAL')).Enabled     := TCheckBox(GetControl('SUPPEXE')).Checked;
    TLabel(GetControl('TDATEECR')).Enabled     := TCheckBox(GetControl('SUPPEXE')).Checked;
    TLabel(GetControl('TDATEECR1')).Enabled    := TCheckBox(GetControl('SUPPEXE')).Checked;
    if (Nature = 'Journal') or ((Nature = 'Balance') or (Nature = 'EXE') and ((Origine= 'SI') or (Origine = 'Sisco II'))) then
    begin
      SetControlChecked('SUPPRTOTAL',False);
      SetControlEnabled('SUPPRTOTAL',False);
    end;
    if TCheckBox(GetControl('SUPPEXE')).Checked then
         THVALComboBox(GetControl('EXERCICE')).ItemIndex :=  THVALComboBox(GetControl('EXERCICE')).Items.count-1
    else
         THVALComboBox(GetControl('EXERCICE')).ItemIndex :=  0;
    ExerciceOnChange (Ecran);
    if (Nature = 'EXE') and not TCheckBox(GetControl('SUPPEXE')).Checked then
    begin
          TCheckBox(GetControl('SUPPEXE')).Checked   := TRUE;
          SetControlEnabled('SUPPEXE',False);
    end;
    if TCheckBox(GetControl('SUPPEXE')).Checked and (GetParamSocSecur ('SO_CPLIENGAMME','') = 'S1') and
   (Origine = 'S1') and ((Nature = 'Journal') or (Nature = 'Balance')) then
            SetControlEnabled ('SUPPANOUVEAU', TRUE)
    else
            SetControlEnabled ('SUPPANOUVEAU', FALSE);
end;

Procedure TOF_ImportConf.ExerciceOnChange(Sender :TObject) ;
var
exo : string;
BEGIN
     exo := THVALComboBox(GetControl('EXERCICE')).Value;
     ExoToDates(exo,
     THEdit(GetControl('DATEECR')),THEdit(GetControl('DATEECR1'))) ;
//GP     if (THVALComboBox(GetControl('EXERCICE')).Text = '<<Tous>>') then
     if (THVALComboBox(GetControl('EXERCICE')).Text = Traduirememoire('<<Tous>>')) then
     begin
          SetControlText ('DATEECR', stDate1900);
          SetControlText ('DATEECR1', stDate2099);
     end;
END ;

Procedure TOF_ImportConf.AnoOnClick(Sender :TObject) ;
var
Q : TQuery;
BEGIN
    THMultiValComboBox(GetControl('JOURNAUX')).Items.Clear;
     THMultiValComboBox(GetControl('JOURNAUX')).Values.Clear;

    if (Origine <> 'S1') or (TCheckBox(GetControl('SUPPANOUVEAU')).Checked = FALSE) then
    Q := OpenSQL('select J_JOURNAL,J_LIBELLE,J_NATUREJAL FROM JOURNAL '+
    'Where J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"CLO" ORDER BY J_JOURNAL' ,true)
    else
    Q := OpenSQL('select J_JOURNAL,J_LIBELLE,J_NATUREJAL FROM JOURNAL ORDER BY J_JOURNAL' ,true);

    while not Q.EOF do
    begin
         THMultiValComboBox(GetControl('JOURNAUX')).Items.Add(Q.FindField('J_LIBELLE').AsString);
         THMultiValComboBox(GetControl('JOURNAUX')).Values.add(Q.FindField('J_JOURNAL').AsString);
         Q.Next;
    end;
    Ferme(Q);
END ;

Initialization
RegisterClasses([TOF_ImportConf]);

end.
