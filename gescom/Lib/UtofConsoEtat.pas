unit UtofConsoEtat;


interface

uses  StdCtrls,Controls,Classes,
      ComCtrls,
      graphics,
      HCtrls,HEnt1,HMsgBox,UTOF, mul, DBGrids,UTOM,Fiche, AglInit,Dialogs,
{$IFDEF EAGLCLIENT}
			eQrs1,
{$ELSE}
			db,forms,sysutils,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      QRS1,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}
{$ENDIF}      
      M3FP, EntGC, FE_main, grids ;

Function  AppelConsostock : boolean ;

Type
     Tof_ConsoEtat = Class (TOF)

     private
        FCBJour, FCBMois, FCBAnnee : THValComboBox;
        DateMin : TDateTime;
        procedure RefreshYearz;
        procedure RefreshMonths;
        procedure RefreshDayz;

     public
        //procedure OnUpdate ; override ;
        procedure OnLoad ; override ;
        procedure OnNew ; override ;
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnYearChange(Sender : TObject);
        procedure OnMonthChange(Sender : TObject);
     END ;

Const
	// libellés des messages
	TexteMessage: array[1..3] of string    = (
           {1}  'La date doit être une fin de semaine (8,15,22 ou fin mois)'
           {2} ,'La date doit être une fin de quinzaine (15 ou fin mois)'
           {3} ,'La date doit être une fin de mois (28,29,30 ou 31)'
               );

implementation

Function AppelConsostock : boolean ;
BEGIN
Result:=False;
if VH_GC.GCDateClotureStock>iDate1900 then
    begin
    Result:=true;
    AGLLanceFiche('GC','GCCONSOCOMP','','','')
    end else
    begin
    PgiBox('Pour accéder à cette fonctionnalité, vous devez avoir effectué#13un traitement de fin de mois',
           'Etat des Consommations');
    end ;
END;

procedure TOF_ConsoEtat.RefreshYearz;
var Q : TQuery;
    YMax,YMin : String;
    i_ind : integer;
begin
FCBAnnee.Clear;
YMax := FormatDateTime('yyyy', VH_GC.GCDateClotureStock);
DateMin:=VH_GC.GCDateClotureStock;

Q := OpenSQL('SELECT Min(GQ_DATECLOTURE) as DateCloture FROM DISPO WHERE GQ_CLOTURE="X" ', true);
if not Q.EOF then DateMin := Q.FindField('DateCloture').AsDateTime;
Ferme(Q);

YMin:=FormatDateTime('yyyy', DateMin);
for i_ind:=StrToInt(YMax) Downto StrToInt(YMin) do FCBAnnee.Items.AddObject(IntToStr(i_ind), Pointer(i_ind));


{FCBAnnee.Items.AddObject(Y, Pointer(StrToInt(Y)));
Q := OpenSQL('SELECT DISTINCT GQ_DATECLOTURE FROM DISPO WHERE GQ_CLOTURE="X" '+
             'ORDER BY GQ_DATECLOTURE DESC', true);
while not Q.EOF do
  begin
  Y := FormatDateTime('yyyy', Q.FindField('GQ_DATECLOTURE').AsDateTime);
  if FCBAnnee.Items.IndexOf(Y) = -1
   then FCBAnnee.Items.AddObject(Y, Pointer(StrToInt(Y)));

  Q.Next;
  end;
Ferme(Q);}

FCBAnnee.ItemIndex := 0;
end;

procedure TOF_ConsoEtat.RefreshMonths;
var //Q : TQuery;
    //D : Tdate;
    Inf, Sup : TDate;
    MoisMax, MoisMin : string;
    i_ind : integer;
begin
FCBMois.Clear;
if FCBAnnee.Items.Count > 0 then
begin
  Inf := EncodeDate(Integer(FCBAnnee.Items.Objects[FCBAnnee.ItemIndex]), 1, 1);
  Sup := EncodeDate(Integer(FCBAnnee.Items.Objects[FCBAnnee.ItemIndex]), 12, 31);
  if Sup>VH_GC.GCDateClotureStock then Sup:=VH_GC.GCDateClotureStock;
  if Inf<DateMin then Inf:=DateMin;
  MoisMax:=FormatDateTime('m', Sup);
  MoisMin:=FormatDateTime('m', Inf);
  for i_ind:=StrToInt(MoisMax) Downto StrToInt(MoisMin) do
      FCBMois.Items.AddObject(LongMonthNames[i_ind], Pointer(i_ind));
end;

{if (FCBAnnee.Text = FormatDateTime('yyyy', VH_GC.GCDateClotureStock))
 then FCBMois.Items.AddObject(FormatDateTime('mmmm', VH_GC.GCDateClotureStock), Pointer(StrToInt(FormatDateTime('m', VH_GC.GCDateClotureStock))));

Inf := EncodeDate(Integer(FCBAnnee.Items.Objects[FCBAnnee.ItemIndex]), 1, 1);
Sup := EncodeDate(Integer(FCBAnnee.Items.Objects[FCBAnnee.ItemIndex]), 12, 31);
Q := OpenSQL('SELECT DISTINCT GQ_DATECLOTURE FROM DISPO WHERE GQ_CLOTURE="X" '+
             'AND GQ_DATECLOTURE>="'+USDateTime(Inf)+'" AND GQ_DATECLOTURE<="'+USDateTime(Sup)+'" '+
             'ORDER BY GQ_DATECLOTURE DESC', true);
while not Q.EOF do
  begin
  D := Q.FindField('GQ_DATECLOTURE').AsDateTime;
  if (FCBMois.Items.IndexOf(FormatDateTime('mmmm', D)) = -1)
   then FCBMois.Items.AddObject(FormatDateTime('mmmm', D), Pointer(StrToInt(FormatDateTime('m', D))));

  Q.Next;
  end;
Ferme(Q);}

FCBMois.ItemIndex := 0;
end;

procedure TOF_ConsoEtat.RefreshDayz;
var Q : TQuery;
    D, Inf, Sup : TDate;
begin
FCBJour.Clear;
if (FCBAnnee.Items.Count > 0) and (FCBMois.Items.Count > 0) then
begin
  if (FCBAnnee.Text = FormatDateTime('yyyy', VH_GC.GCDateClotureStock)) and (FCBMois.Text = FormatDateTime('mmmm', VH_GC.GCDateClotureStock))
   then FCBJour.Items.AddObject(FormatDateTime('dd', VH_GC.GCDateClotureStock), Pointer(StrToInt(FormatDateTime('d', VH_GC.GCDateClotureStock))));

  Inf := EncodeDate(Integer(FCBAnnee.Items.Objects[FCBAnnee.ItemIndex]),
                    Integer(FCBMois.Items.Objects[FCBMois.ItemIndex]), 1);
  Sup := FinDeMois(Inf);
  Q := OpenSQL('SELECT DISTINCT GQ_DATECLOTURE FROM DISPO WHERE GQ_CLOTURE="X" '+
               'AND GQ_DATECLOTURE>="'+USDateTime(Inf)+'" AND GQ_DATECLOTURE<="'+USDateTime(Sup)+'" '+
               'ORDER BY GQ_DATECLOTURE DESC', true);
  while not Q.EOF do
    begin
    D := Q.FindField('GQ_DATECLOTURE').AsDateTime;
    if (FCBJour.Items.IndexOf(FormatDateTime('dd', D)) = -1)
     then FCBJour.Items.AddObject(FormatDateTime('dd', D), Pointer(StrToInt(FormatDateTime('d', D))));

    Q.Next;
    end;
  Ferme(Q);

  FCBJour.ItemIndex := 0;
end;
end;

procedure TOF_ConsoEtat.OnYearChange(Sender : TObject);
begin
RefreshMonths;
RefreshDayz;
end;

procedure TOF_ConsoEtat.OnMonthChange(Sender : TObject);
begin
RefreshDayz;
end;

// Initialisation Combos de sélection
procedure TOF_ConsoEtat.OnArgument (stArgument : String ) ;
begin
  Inherited;
// paramétrage du libellé etablissement quand on est en multi-dépôts
if VH_GC.GCMultiDepots then
   SetControlText('TGQ_DEPOT',TraduireMemoire('Dépôts'))
else
   SetControlText('TGQ_DEPOT',TraduireMemoire('Etablissement'));

THValComboBox (Ecran.FindComponent('GQ_DEPOT')).ItemIndex := 0;
THValComboBox (Ecran.FindComponent('GA_FAMILLENIV1')).ItemIndex := 0;
THValComboBox (Ecran.FindComponent('GA_FAMILLENIV2')).ItemIndex := 0;
THValComboBox (Ecran.FindComponent('GA_FAMILLENIV3')).ItemIndex := 0;
THValComboBox (Ecran.FindComponent('RUPT1')).ItemIndex := 0;
end;

// Positionnement Var Société "Période clôture Stocks" et "Libellés Familles"
procedure TOF_ConsoEtat.OnNew;
var st_Etiq : string;
begin
  Inherited;
FCBJour := THValComboBox(Ecran.FindComponent('CBJOUR'));
FCBMois := THValComboBox(Ecran.FindComponent('CBMOIS'));
FCBAnnee := THValComboBox(Ecran.FindComponent('CBANNEE'));
FCBAnnee.OnChange := OnYearChange;
FCBMois.OnChange := OnMonthChange;

RefreshYearz;
RefreshMonths;
RefreshDayz;

SetControlText('REFPERIODE', VH_GC.GCPeriodeStock);
SetControlText('DATMAX', DateToStr(VH_GC.GCDateClotureStock));
St_Etiq := RechDom ('GCLIBFAMILLE', 'LF1', True);
if St_Etiq<>'' then
    begin
    SetControlText('TGA_FAMILLENIV1', St_Etiq);
    end else
    begin
    SetControlText('TGA_FAMILLENIV1', '');
    SetControlVisible('TGA_FAMILLENIV1', false);
    SetControlEnabled('GA_FAMILLENIV1', false);
    SetControlVisible('GA_FAMILLENIV1', false);
    end;
St_Etiq := RechDom ('GCLIBFAMILLE', 'LF2', True);
if St_Etiq<>'' then
    begin
    SetControlText('TGA_FAMILLENIV2', St_Etiq);
    end else
    begin
    SetControlText('TGA_FAMILLENIV2', '');
    SetControlVisible('TGA_FAMILLENIV2', false);
    SetControlEnabled('GA_FAMILLENIV2', false);
    SetControlVisible('GA_FAMILLENIV2', false);
    end;
St_Etiq := RechDom ('GCLIBFAMILLE', 'LF3', True);
if St_Etiq<>'' then
    begin
    SetControlText('TGA_FAMILLENIV3', St_Etiq);
    end else
    begin
    SetControlText('TGA_FAMILLENIV3', '');
    SetControlVisible('TGA_FAMILLENIV3', false);
    SetControlEnabled('GA_FAMILLENIV3', false);
    SetControlVisible('GA_FAMILLENIV3', false);
    end;
end;

 // Préparation des 12 périodes et Génération SQL pour Tri
procedure TOF_ConsoEtat.OnLoad;
Var i_ind : integer;
    St_Order : string ;
    Wdat1, Wdat2, DateConso : TdateTime;
    Waa, Wmm, Wjj : word;
begin
   Inherited;
DateConso := EncodeDate(Integer(FCBAnnee.Items.Objects[FCBAnnee.ItemIndex]),
                        Integer(FCBMois.Items.Objects[FCBMois.ItemIndex]),
                        Integer(FCBJour.Items.Objects[FCBJour.ItemIndex]));
SetControlText('DATMAX', DateToStr(DateConso));
// Contrôle date saisie
Wdat1 := DateConso;
DecodeDate(Wdat1,Waa,Wmm,Wjj);
Wdat2 := FinDeMois(Wdat1);
If ((GetControlText('REFPERIODE')='MOI') and (Wdat1<>Wdat2)) or
   ((GetControlText('REFPERIODE')='QUI') and (Wdat1<>Wdat2) and (Wjj<>15)) or
   ((GetControlText('REFPERIODE')='SEM') and (Wdat1<>Wdat2) and (Wjj<>08) and (Wjj<>15) and (Wjj<>22))
   then
   begin
     Waa:=1900;
     Wmm:=01;
     Wjj:=31;
     Wdat1 := EncodeDate(Waa,Wmm,Wjj);
     SetControlText('DATMAX',DateToStr(Wdat1));
     If (GetControlText('REFPERIODE')='MOI') then
        begin
          LastError:=3;
          LastErrorMsg:=TexteMessage[LastError];
        end;
     If (GetControlText('REFPERIODE')='QUI') then
        begin
          LastError:=2;
          LastErrorMsg:=TexteMessage[LastError];
        end;
     If (GetControlText('REFPERIODE')='SEM') then
        begin
          LastError:=1;
          LastErrorMsg:=TexteMessage[LastError];
        end;
     exit;
   end;
Wdat2 := Wdat1;
// Calcul des 12 périodes mensuelles
If GetControlText('REFPERIODE')='MOI'
   then
   begin
     SetControlText('LIBPERIODE','Mensuelle');
     For i_ind := 12 downto 1 do
     begin
       SetControlText('DATE'+InttoStr(i_ind),DateToStr(Wdat2));
       Wdat1 := PlusDate(Wdat2,-1,'M');
       Wdat2 := FinDeMois(Wdat1);
     end;
   end;
// calcul des 12 Périodes quinzaines
If GetControlText('REFPERIODE')='QUI'
   then
   begin
     SetControlText('LIBPERIODE','à la quinzaine');
     For i_ind := 12 downto 1 do
     begin
       SetControlText('DATE'+InttoStr(i_ind),DateToStr(Wdat2));
       DecodeDate(Wdat2,Waa,Wmm,Wjj);
       if Wjj > 15
       then
         begin
           Wjj := 15;
           Wdat2 := EncodeDate(Waa,Wmm,Wjj);
         end
       else
         begin
           Wdat1 := PlusDate(wdat2,-1,'M');
           Wdat2 := FinDeMois(wdat1);
         end;
     end;
   end;
// calcul des 12 Périodes hebdo
If GetControlText('REFPERIODE')='SEM'
   then
   begin
     SetControlText('LIBPERIODE','hebdomadaire');
     For i_ind := 12 downto 1 do
     begin
       SetControlText('DATE'+InttoStr(i_ind),DateToStr(Wdat2));
       DecodeDate(Wdat2,Waa,Wmm,Wjj);
       Case Wjj of
       22: begin
           Wjj := 15;
           Wdat2 := EncodeDate(Waa,Wmm,Wjj);
           end;
       15: begin
           Wjj := 08;
           Wdat2 := EncodeDate(Waa,Wmm,Wjj);
           end;
       08: begin
           Wdat1 := PlusDate(wdat2,-1,'M');
           Wdat2 := FinDeMois(wdat1);
           end;
       else
           begin
           Wjj := 22;
           Wdat2 := EncodeDate(Waa,Wmm,Wjj);
           end;
     end;
   end;
end;
// Formatage des 12 colonnes "dates"
for i_ind :=1 to 12 do
begin
     Wdat1 := StrToDate(GetControlText('DATE'+InttoStr(i_ind)));
     SetControlText('LIBDAT'+InttoStr(i_ind),FormatDateTime('dd/mm',Wdat1));
end;
// Préparation du Tri complémentaire
St_Order:='';
//For i_ind := 1 to 3 do
//begin
//     St := string (THEdit (Ecran.FindComponent('XX_RUPTURE'+InttoStr(i_ind))).Text);
//     If St = '' then Break;
//     St_Order:=St_Order+St+',';
//end;
SetControlText('XX_ORDERBY', St_Order + 'GQ_ARTICLE');
// Forçage des ruptures de pages majeures
//For i_ind := 3 downto 2 do
//begin
     //if TCheckBox (Ecran.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Checked = true
     //then TCheckBox (Ecran.FindComponent('SAUTRUPT'+InttoStr(i_ind - 1))).Checked := true;
//end;
end;

///////// M.à.j. des lignes du ComboBox "Ruptures" //////////
Procedure TOF_ConsoEtat_AffectGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value, St_Text, St : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCCONSOCOMP') then exit;
Indice := Integer (Parms[1]);
if VH_GC.GCMultiDepots then St_Plus := ' AND CO_CODE<>"ETA" '
else St_Plus := ' AND CO_CODE<>"DEP" ';
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
St_Text := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
For i_ind := 1 to 3 do
    BEGIN
    if i_ind = Indice then continue;
    St := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value);
    If St <> '' then St_Plus := St_Plus + ' AND CO_CODE <>"' + St + '"';
    END;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Plus := St_Plus;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text := St_Text;
if St_value <>'' then THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value := St_Value
                 else THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).ItemIndex := 0;
END;

///////// (Dé)Activation lignes de Ruptures déjà sélectionnées //////////
Procedure TOF_ConsoEtat_ChangeGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCCONSOCOMP') then exit;
Indice := Integer (Parms[1]);
if VH_GC.GCMultiDepots then St_Plus := ' AND CO_CODE<>"ETA" '
else St_Plus := ' AND CO_CODE<>"DEP" ';
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
if St_Value = '' then
    BEGIN
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text := '';
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text := '';
    THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Plus := St_Plus;
    //THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text := '';
    TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Checked := False;
    TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Enabled := False;
    For i_ind := Indice + 1 to 3 do
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value := '';
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Plus := St_Plus;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := False;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Color := clBtnFace;
        THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(i_ind))).Text := '';
        THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(i_ind))).Text := '';
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Checked := False;
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Enabled := False;
        END;
    END else
    BEGIN
    if Indice < 3 then
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Color := clWindow;
        //TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice + 1))).Enabled := True;
        END;
    TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Enabled := True;
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text :=
            RechDom ('GCGROUPCONSO', St_Value, True);
    if (St_value = 'LF1') or (St_value = 'LF2') or (St_value = 'LF3') then
       begin
         THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text :=
           RechDom ('GCLIBFAMILLE', St_value, true);
       end else
       begin
         THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text :=
           string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
       end;
    END;
END;

///////// Activation des sauts de page sur niveaux supérieurs /////////
Procedure TOF_ConsoEtat_ChangeSautPage (parms:array of variant; nb: integer ) ;
var F: TForm;
    i_ind, Indice : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCCONSOCOMP') then exit;
Indice := Integer (Parms[1]);
if TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Checked then
    BEGIN
    For i_ind := 1 to Indice - 1 do
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Checked := True;
    END else
    BEGIN
    For i_ind := Indice + 1 to 6 do
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Checked := False;
    END;
END;
/////////////////////////////////////////////////////////////////////////////

procedure InitTOFConsoEtat ();
begin
RegisterAglProc('ConsoEtat_ChangeGroup', True , 1, TOF_ConsoEtat_ChangeGroup);
RegisterAglProc('ConsoEtat_AffectGroup', True , 1, TOF_ConsoEtat_AffectGroup);
RegisterAglProc('ConsoEtat_ChangeSautPage', True , 0, TOF_ConsoEtat_ChangeSautPage);
end;

Initialization
registerclasses([Tof_ConsoEtat]);
InitTOFConsoEtat();
end.
