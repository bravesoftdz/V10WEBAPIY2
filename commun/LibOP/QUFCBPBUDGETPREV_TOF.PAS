{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/10/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFCBPBUDGETPREV ()
Mots clefs ... : TOF;QUFCBPBUDGETPREV
*****************************************************************}
Unit QUFCBPBUDGETPREV_TOF ;

Interface

Uses Classes,
     sysutils,
     HCtrls,
     HMsgBox,
     hEnt1,
     UTOF,
     UTOB,
     ed_tools,
     paramsoc,
{$IFNDEF EAGLCLIENT}    //CEGID-CCMX le 07/11/2006 DEBUT
    {$IFNDEF DBXPRESS}dbtables {BDE}, {$ELSE}uDbxDataSet, {$ENDIF}
     FE_main,
    db,
{$else}
		 Maineagl,
{$ENDIF}               //CEGID-CCMX le 07/11/2006 FIN
     windows,
     Controls,
     htb97;

Type
  TOF_QUFCBPBUDGETPREV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DetailAuJour(VertFields, HorzFields, CellValue: hString);
   private
    codesession,codeFiltre:hString;
    sTypeDA : string;
    {$IFNDEF MODE}{$IFDEF GCGC} sSessionN, sSessionN1 : string;  {$ENDIF}{$ENDIF} //CEGID-CCMX le 07/11/2006
    { EVI / Date Edition Debut, Date Edition Fin }
    date1,date2,DateEditionDeb,DateEditionFin:TDateTime;
    TabLibelle:array of hString;
    TabHisto : array[ 0..6 ] of string;
    Incr:integer;
    InProgress : boolean;
    BCherche: TToolbarButton97;
    procedure MyOnAfterFormShow;
    procedure RemplitTableQBPCubeTmpPgi0(Restrictions:hString);
    procedure DonneFiltre;
    procedure ValAfficheOnChange(Sender: TObject);
    procedure ChangePres(Libelle,FormuleSQL1, FormuleSQL2, FormuleSQL3:string);
    procedure BChercheOnClick(Sender: TObject);
    procedure ModifTOB(T : TOB;ValAff:integer);
    procedure InitChamp(ValAff:integer);
    function VerifPres(ValAff:integer) : boolean;
    procedure VerifValeurs(Sender:TObject);

end ;

const
   TabPrevu : array[ 0..6 ] of string = (
    ( 'QBQ_PREVU' ),
    ( 'QBQ_CAPREVU' ),
    ( 'QBQ_CAPREVU2' ),
    ( 'QBQ_CAPREVU3' ),
    ( 'QBQ_CAPREVU4' ),
    ( 'QBQ_CAPREVU5' ),
    ( 'QBQ_CAPREVU6' )
    );

   TabRealise : array[ 0..6 ] of string = (
    ( 'QBQ_REALISE' ),
    ( 'QBQ_CAREALISE' ),
    ( 'QBQ_CAREALISE2' ),
    ( 'QBQ_CAREALISE3' ),
    ( 'QBQ_CAREALISE4' ),
    ( 'QBQ_CAREALISE5' ),
    ( 'QBQ_CAREALISE6' )
    );

Implementation

uses cube, UUtil, BPCubeTmp, BPBasic, BPFctSession, SynScriptBP,BPUtil,
  TntForms,StrUtils;

procedure TOF_QUFCBPBUDGETPREV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFCBPBUDGETPREV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFCBPBUDGETPREV.OnUpdate ;
begin
  Inherited ;
  //filtre
  DonneFiltre;
end ;


procedure TOF_QUFCBPBUDGETPREV.RemplitTableQBPCubeTmpPgi0(Restrictions:hString);
var nivMax,i:integer;
    TabCodeAxe:array [1..11] of hString;
    codeAxe:hString;
begin
 { EVI / Ajout Date Edition Debut, Date Edition Fin }
 { EVI / Génération du cube avec restriction utilisateur }
 RemplitTableQBPCubeTmpPgi(codeSession,Restrictions,date1,date2,DateEditionDeb,DateEditionFin);

 ChercheCodeAxeSession(codeSession,TabCodeAxe);
 if not SessionInitPrev(codesession)
  then nivMax:=ChercheNivMax(codeSession)
  else
   begin
    nivMax:=ChercheNivMaxSession(codeSession);
    TabCodeAxe[NivMax+1]:='ARTICLE';
   end;
 TFCube(Ecran).lDispo.Clear;
 TFCube(Ecran).lDispo.Refresh;
 for i:=1 to nivMax do
  begin
   codeAxe:=DonneLibelleCAxe('','',TabCodeAxe[i+1]);
   Inc(Incr);
   TFCube(Ecran).addFormule(codeAxe,'QBQ_VALAXECT'+IntToStr(i),'AXE'+IntToStr(Incr));
   TFCube(Ecran).addFormule(TraduireMemoire('Libellé de ') + codeAxe, 'QBQ_LIBVALAXECT'+IntToStr(i),'LIBAXE'+IntToStr(Incr));
  end;
 if ContextBp = 3 then TFCube(Ecran).addFormule(TraduireMemoire('Type de Valeur'), 'QBQ_NATVALAFF','QBQ_NATVALAFF');
end;

procedure TOF_QUFCBPBUDGETPREV.DonneFiltre;
var i,Dot,DotTemp:integer;
AxeTemp:hString;
begin
  codeFiltre:='';
  Dot :=0;
  codeFiltre:=TFCube(ecran).wherecritere;

  { EVI / Récupération des critères - multiselection}
  codefiltre:=' QBQ_CODESESSION="'+codesession+'" ';
  for i:=1 to 10 do
  begin
    if (GetControlText('QBQ_VALAXECT'+intToStr(i))<>'') and (Pos(';',GetControlText('QBQ_VALAXECT'+intToStr(i)))<>0)
    then
    begin
      Dot := VerifPos(true,GetControlText('QBQ_VALAXECT'+intToStr(i)), ';',Dot);
      AxeTemp := Copy(GetControlText('QBQ_VALAXECT'+intToStr(i)),1,Dot-2);
      codefiltre:= codefiltre+' AND QBQ_VALAXECT'+intToStr(i)+' in ("'+AxeTemp+'"';
      while Dot <> -1 do
      begin
        if Dot <> -1 then
        begin
          DotTemp := VerifPos( true, GetControlText('QBQ_VALAXECT'+intToStr(i)), ';', Dot );
          if DotTemp <> -1 then AxeTemp := Copy( GetControlText('QBQ_VALAXECT'+intToStr(i)), Dot,DotTemp-Dot-1)
          else AxeTemp := Copy( GetControlText('QBQ_VALAXECT'+intToStr(i)), Dot,length(GetControlText('QBQ_VALAXECT'+intToStr(i))));
          if LastDelimiter(';',AxeTemp) = length(AxeTemp) then AxeTemp := copy(AxeTemp,1,length(AxeTemp)-1);
          codefiltre:= codefiltre+',"'+AxeTemp+'"';
          Dot := DotTemp;
        end
      end;
    codefiltre:= codefiltre+')';
    end
    else if GetControlText('QBQ_VALAXECT'+intToStr(i))<>'' then
    if Uppercase(GetControlText('QBQ_VALAXECT'+intToStr(i)))<>Uppercase(TraduireMemoire('<<Tous>>')) then codefiltre:=codefiltre+' AND QBQ_VALAXECT'+intToStr(i)+' like "'+GetControlText('QBQ_VALAXECT'+intToStr(i))+'%" ';
  end;

  TFCube(ecran).wherecritere:=codeFiltre;
end;

procedure TOF_QUFCBPBUDGETPREV.OnLoad ;
var RI,i,nivMax:integer;
ValueAxe,LimitCube : hString;
Q:TQuery;
begin
  Inherited ;
  blocageMenu(TFCube(ecran));
  LimitCube := '';

  { EVI / Contrôle pour le cas où la tablette est vide (session non initialisée) }
  if (GetControlText('DATEDEB')='') or (GetControlText('DATEFIN')='')
    then
   begin
     DateEditionDeb := 0;
     DateEditionFin := 0;
   end
   else
   begin
     DateEditionDeb := StrToDate(GetControlText('DATEDEB'));
     DateEditionFin := StrToDate(GetControlText('DATEFIN'));
   end;

  codesession := GetControlText('QBQ_CODESESSION');
  //lit le code session et le vérifie
  //si il est incorrect ou non renseigné on ne fait rien
  if (ControleSession(codesession)=false)
  then
   begin
    PGIINFO('Le code session doit être renseigné correctement.');
    deblocageMenu(TFCube(ecran));
    exit;
   end;
  //CEGID-CCMX le 07/11/2006 DEBUT
  {$IFNDEF MODE}
  {$IFDEF GCGC}
  sTypeDA := '';
  sSessionN := GetParamsocSecur('SO_GCCODESESSIONN','');
  sSessionN1 := GetParamsocSecur('SO_GCCODESESSIONN1','');
  // le 16/01/2007 : en fait on limite pas seulement au session N ou N+1 mais aux
  // sessions utilisées
  if (sSessionN = codesession) or (sSessionN1 = codesession)
  or (ExisteSQL('SELECT DA_CODESESSION FROM PIECEDA WHERE DA_CODESESSION="' + codesession+'"'))
  then sTypeDA:='TYPEDA';
  {$ENDIF}
  {$ENDIF}
  //CEGID-CCMX le 07/11/2006 FIN

  date1 :=StrToDateTime(GetControlText('EDTDATE1'));
  date2 :=StrToDateTime(GetControlText('EDTDATE2'));

  nivMax:=ChercheNivMaxSession(codeSession);
  //Récupère les restrictions pour la génération du cube
  for i:=1 to nivMax do
  begin
     ValueAxe := GetControlText('QBQ_VALAXECT'+IntToStr(i));
     ValueAxe:= AnsiReplaceText(ValueAxe, ';', '","');
     if LimitCube = '' then LimitCube := '("'+ValueAxe+'")'
     else LimitCube := LimitCube + '@@' + '("'+ValueAxe+'")'
  end;

  { EVI / Suppression axes, formules des sessions précédentes }
  //AJOUT SUPRESSION LIBELLE
  for i:=1 to Incr do
  begin
    TFCube(Ecran).delFormule('AXE'+IntToStr(i));
    TFCube(Ecran).delFormule('LIBAXE'+IntToStr(i));
  end;
  if ContextBP = 3 then TFCube(Ecran).delFormule('QBQ_NATVALAFF');

  Incr:=0;

  RI:=DonneValeurAffiche(codeSession);
  Case ContextBP of
    0,1 : begin //Mode-GC
          case RI of
            1 : THValComboBox(GetControl('ValAffiche')).Value := 'PTC';
            2 : THValComboBox(GetControl('ValAffiche')).Value := 'QTE';
            3 : THValComboBox(GetControl('ValAffiche')).Value := 'PHT';
            4 : THValComboBox(GetControl('ValAffiche')).Value := 'UHT';
            5 : THValComboBox(GetControl('ValAffiche')).Value := 'UTC';
            6 : THValComboBox(GetControl('ValAffiche')).Value := 'PAH';
            7 : THValComboBox(GetControl('ValAffiche')).Value := 'MAR';
          end;
        end;
    2 : begin //compta
          case RI of
            1 : THValComboBox(GetControl('ValAffiche')).Value := 'DC1';
            3 : THValComboBox(GetControl('ValAffiche')).Value := 'CD1';
          end
        end;
    3 : begin //Paie
          LibValAff(codesession,TabLibelle);
          THValComboBox(GetControl('ValAffiche')).OnChange(self);
        end
  end;

  InitMoveProgressForm(nil, 'Cube', TraduireMemoire('Veuillez patienter'), NbDeMaille(codesession,'InitOK',DateEditionDeb, DateEditionFin) + NbMajCoeff(codesession), True, True);
  InProgress := true;
  Ferme(Q);

  RemplitTableQBPCubeTmpPgi0(LimitCube);

  TFCube(ecran).ForceSelect('',TFCube(ecran).FDimensions.Value);

  deblocageMenu(TFCube(ecran));

  FiniMoveProgressForm;
  InProgress := false;
end ;

procedure TOF_QUFCBPBUDGETPREV.OnArgument (S : String ) ;
begin
  Inherited ;

  TFCube(ecran).OnAfterFormShow:=MyOnAfterFormShow;
  BCherche := GetControl('BCherche') as TToolbarButton97;
  if ContextBP = 2 then THValComboBox(GetControl('ValAffiche')).DataType := 'QUTBPANALYTIQVALAFF'
  else THValComboBox(GetControl('ValAffiche')).DataType := 'QUTBPVUEARBRE001';
  THValComboBox(GetControl('ValAffiche')).OnChange := ValAfficheOnChange;
  TToolBarButton97(GetControl('BCherche')).OnClick := BChercheOnClick;
  { EVI / Trap de la touche ESC : éviter de sortir pendant la génération }
  Ecran.OnKeyDown := FormKeyDown;
  InProgress := false;
  { EVI / Double clic sur le cube pour le détail au jour }
  if ContextBP <> 3 then TFCube(Ecran).OnListeDblClick := DetailAuJour;

  TFCube(Ecran).lValeurs.OnEnter := VerifValeurs;

  Case ContextBP of
    0,1 : begin //Mode-GC
          TabHisto[0] := 'QBQ_HISTO';
          TabHisto[1] := 'QBQ_CAHISTO';
          TabHisto[2] := 'QBQ_CAHISTO2';
          TabHisto[3] := 'QBQ_CAHISTO3';
          TabHisto[4] := 'QBQ_CAHISTO4';
          TabHisto[5] := 'QBQ_CAHISTO5';
          TabHisto[6] := 'QBQ_CAHISTO6';

          SetLength(TabLibelle,7);
          TabLibelle[0]:=TraduireMemoire('Quantité');
          TabLibelle[1]:=TraduireMemoire('CA TTC net');
          TabLibelle[2]:=TraduireMemoire('CA HT net');
          TabLibelle[3]:=TraduireMemoire('CA HT avant remise');
          TabLibelle[4]:=TraduireMemoire('CA TTC avant remise');
          TabLibelle[5]:=TraduireMemoire('Cumul achat en PAMP');
          TabLibelle[6]:=TraduireMemoire('Marge');
        end;
    2 : begin //compta
          TabHisto[0] := 'QBQ_HISTO';
          TabHisto[1] := 'QBQ_CAHISTO';
          TabHisto[2] := 'QBQ_CAHISTO2';
          TabHisto[3] := 'QBQ_CAHISTO3';
          TabHisto[4] := 'QBQ_CAHISTO4';
          TabHisto[5] := 'QBQ_CAHISTO5';
          TabHisto[6] := 'QBQ_CAHISTO6';
          SetLength(TabLibelle,3);
          TabLibelle[1]:=TraduireMemoire('Débit - Crédit');
          TabLibelle[2]:=TraduireMemoire('Crédit - Débit');
        end;
    3 : begin //Paie
          TabHisto[0] := 'QBQ_VALAFF2';
          TabHisto[1] := 'QBQ_VALAFF1';
          TabHisto[2] := 'QBQ_VALAFF3';
          TabHisto[3] := 'QBQ_VALAFF4';
          TabHisto[4] := 'QBQ_VALAFF5';
          TabHisto[5] := 'QBQ_VALAFF6';
          TabHisto[6] := 'QBQ_VALAFF7';
          SetLength(TabLibelle,8);
        end;
  end;
end ;

procedure TOF_QUFCBPBUDGETPREV.ModifTOB(T : TOB;ValAff:integer);
var i : integer;

  function ReplaceChamp(chp:string):string;
  var i : integer;
  begin
    Result:=chp;
    for i:=0 to 6 do
    begin
      if ((i <> ValAff) AND (i <> 1))  then
      begin
        Result:=StringReplace(Result, TabHisto[i], TabHisto[ValAff],[rfReplaceAll, rfIgnoreCase]);
        Result:=StringReplace(Result, TabPrevu[i], TabPrevu[ValAff],[rfReplaceAll, rfIgnoreCase]);
        Result:=StringReplace(Result, TabRealise[i], TabRealise[ValAff],[rfReplaceAll, rfIgnoreCase]);
      end;
    end;
    if result=chp then
    begin
       try
         StrToInt(Copy(Result,Pos(TabHisto[1],Result)+length(TabHisto[1]),1))
       except
         on EConvertError do Result:=StringReplace(Result, TabHisto[1], TabHisto[ValAff],[rfReplaceAll, rfIgnoreCase]);
       end;

       try
         StrToInt(Copy(Result,Pos(TabPrevu[1],Result)+length(TabPrevu[1]),1))
       except
         on EConvertError do Result:=StringReplace(Result, TabPrevu[1], TabPrevu[ValAff],[rfReplaceAll, rfIgnoreCase]);

       end;

       try
         StrToInt(Copy(Result,Pos(TabRealise[1],Result)+length(TabRealise[1]),1))
       except
         on EConvertError do Result:=StringReplace(Result, TabRealise[1], TabRealise[ValAff],[rfReplaceAll, rfIgnoreCase]);
       end;

    end;
  end;

begin
  Case ContextBP of
   0,1,2 : begin
           if Pos('FORMULE',T.GetValue('DH_NOMCHAMP')) > 0 then  T.PutValue('FORMULESQL',ReplaceChamp(T.GetValue('FORMULESQL')))
           else
           begin
             if Pos('HISTO',T.GetValue('FORMULESQL')) > 0 then
             begin
               TFCube(Ecran).delFormule(T.GetValue('DH_NOMCHAMP'));
               T.PutValue('FORMULESQL',TabHisto[ValAff]);
               T.PutValue('DH_NOMCHAMP',TabHisto[ValAff]);
               T.PutValue('TITRE',TabLibelle[ValAff]+' Historique');
               TFCube(Ecran).addFormule(TabLibelle[ValAff]+' Historique',TabHisto[ValAff],TabHisto[ValAff]);
             end;

             if Pos('PREVU',T.GetValue('FORMULESQL')) > 0 then
             begin
               TFCube(Ecran).delFormule(T.GetValue('DH_NOMCHAMP'));
               T.PutValue('FORMULESQL',TabPrevu[ValAff]);
               T.PutValue('DH_NOMCHAMP',TabPrevu[ValAff]);
               T.PutValue('TITRE',TabLibelle[ValAff]+' Prévu');
               TFCube(Ecran).addFormule(TabLibelle[ValAff]+' Prévu',TabPrevu[ValAff],TabPrevu[ValAff]);
             end;

             if Pos('REALISE',T.GetValue('FORMULESQL'))> 0 then
             begin
               TFCube(Ecran).delFormule(T.GetValue('DH_NOMCHAMP'));
               T.PutValue('FORMULESQL',TabRealise[ValAff]);
               T.PutValue('DH_NOMCHAMP',TabRealise[ValAff]);
               T.PutValue('TITRE',TabLibelle[ValAff]+' Realisé');
               TFCube(Ecran).addFormule(TabLibelle[ValAff]+' Realisé',TabRealise[ValAff],TabRealise[ValAff]);
             end;
           end;
         end;
   3 : begin
         if Pos('FORMULE',T.GetValue('DH_NOMCHAMP')) > 0 then  T.PutValue('FORMULESQL',ReplaceChamp(T.GetValue('FORMULESQL')))
         else
         begin
           for i := 1 to NbrValAff(codeSession) do
           if Pos(TabHisto[i],T.GetValue('FORMULESQL')) > 0 then
           begin
             TFCube(Ecran).delFormule(T.GetValue('DH_NOMCHAMP'));
             T.PutValue('FORMULESQL',TabHisto[i]);
             T.PutValue('DH_NOMCHAMP',TabHisto[i]);
             T.PutValue('TITRE',TabLibelle[i+1]);
             TFCube(Ecran).addFormule(TabLibelle[i+1],TabHisto[i],TabHisto[i]);
           end;
         end;
       end;

  end;
end;


procedure TOF_QUFCBPBUDGETPREV.InitChamp(ValAff:integer);
var i:integer;
begin
  for i:=0 to 6 do
  begin
    if TFCube(Ecran).ExistFormule(TabHisto[i]) then TFCube(Ecran).delFormule(TabHisto[i]);
    if TFCube(Ecran).ExistFormule(TabPrevu[i]) then TFCube(Ecran).delFormule(TabPrevu[i]);
    if TFCube(Ecran).ExistFormule(TabRealise[i]) then TFCube(Ecran).delFormule(TabRealise[i]);
  end;
  if TFCube(Ecran).ExistFormule('YYY1') then TFCube(Ecran).delFormule('YYY1');
  if TFCube(Ecran).ExistFormule('YYY2') then TFCube(Ecran).delFormule('YYY2');
  //CEGID-CCMX
  if (sTypeDA = 'TYPEDA') then
  begin
    if TFCube(Ecran).ExistFormule('QBQ_CAPREAL2') then TFCube(Ecran).delFormule('QBQ_CAPREAL2');
    if TFCube(Ecran).ExistFormule('YYY3') then TFCube(Ecran).delFormule('YYY3');
  end;

  Case ContextBP of
    0,1,2 : begin
            TFCube(Ecran).addFormule(TabLibelle[ValAff]+' Historique',TabHisto[ValAff],TabHisto[ValAff]);
            TFCube(Ecran).addFormule(TabLibelle[ValAff]+' Prévu',TabPrevu[ValAff],TabPrevu[ValAff]);
            TFCube(Ecran).addFormule(TabLibelle[ValAff]+' Realisé',TabRealise[ValAff],TabRealise[ValAff]);
            TFCube(Ecran).addFormule(TraduireMemoire('Delta Réalisé - Prévu'),TabRealise[ValAff]+'-'+TabPrevu[ValAff],'YYY1');
            TFCube(Ecran).addFormule(TraduireMemoire('Delta Réalisé - Historique'),TabRealise[ValAff]+'-'+TabHisto[ValAff],'YYY2');
            //CEGID-CCMX le 07/11/2006 DEBUT
            if (sTypeDA = 'TYPEDA') then
            begin
              TFCube(Ecran).addFormule(TabLibelle[2] + ' Presque Réalisé','QBQ_CAPREAL2','QBQ_CAPREAL2');
              TFCube(Ecran).addFormule('Delta Total Réalisé - ' + TabLibelle[2]+' Prévu',TabRealise[2]+'+QBQ_CAPREAL2-'+TabPrevu[2],'YYY3');
            end;
            //CEGID-CCMX le 07/11/2006 FIN
          end;
    3 : begin
          for i := 1 to NbrValAff(codeSession) do
          begin
            TFCube(Ecran).addFormule(TabLibelle[i],TabHisto[i-1],TabHisto[i-1]);
          end;
        end;
  end;
end;

function TOF_QUFCBPBUDGETPREV.VerifPres(ValAff:integer) : boolean;
var IndexFChoix : integer;
    i,j,okaxe:integer;
begin
  IndexFChoix := TFCube(ecran).FChoix.ItemIndex;
  TFCube(ecran).FChoix.ItemIndex := 0;
  result:=false;
  okaxe := 0;

  Case ContextBP of
    0,1,2 : begin
            for i := 0 to TFCube(ecran).lDispo.Items.Count-1 do
            begin
              if Pos(TabHisto[ValAff],TFCube(ecran).lDispo.Items.Strings[i]) > 0 then okaxe := okaxe + 1 ;
              if Pos(TabPrevu[ValAff],TFCube(ecran).lDispo.Items.Strings[i]) > 0 then okaxe := okaxe + 1 ;
              if Pos(TabRealise[ValAff],TFCube(ecran).lDispo.Items.Strings[i]) > 0 then okaxe := okaxe + 1 ;
              if (sTypeDA = 'TYPEDA') then
              begin
                if Pos('QBQ_CAPREAL2',TFCube(ecran).lDispo.Items.Strings[i]) > 0 then okaxe := okaxe + 1 ;
              end;
            end;

            for i := 0 to TFCube(ecran).lLigne.Items.Count-1 do
            begin
              if Pos(TabHisto[ValAff],TFCube(ecran).lLigne.Items.Strings[i]) > 0 then okaxe := okaxe + 1 ;
              if Pos(TabPrevu[ValAff],TFCube(ecran).lLigne.Items.Strings[i]) > 0 then okaxe := okaxe + 1 ;
              if Pos(TabRealise[ValAff],TFCube(ecran).lLigne.Items.Strings[i]) > 0 then okaxe := okaxe + 1 ;
              if (sTypeDA = 'TYPEDA') then
              begin
                if Pos('QBQ_CAPREAL2',TFCube(ecran).lLigne.Items.Strings[i]) > 0 then okaxe := okaxe + 1 ;
              end;
            end;

            for i := 0 to TFCube(ecran).lValeurs.Items.Count-1 do
            begin
              if Pos(TabHisto[ValAff],TFCube(ecran).lValeurs.Items.Strings[i]) > 0  then okaxe := okaxe + 1 ;
              if Pos(TabPrevu[ValAff],TFCube(ecran).lValeurs.Items.Strings[i]) > 0  then okaxe := okaxe + 1 ;
              if Pos(TabRealise[ValAff],TFCube(ecran).lValeurs.Items.Strings[i]) > 0  then okaxe := okaxe + 1 ;
              if (sTypeDA = 'TYPEDA') then
              begin
                if Pos('QBQ_CAPREAL2',TFCube(ecran).lValeurs.Items.Strings[i]) > 0 then okaxe := okaxe + 1 ;
              end;
            end;
          end;

    3 : begin
          for j := 1 to NbrValAff(codeSession) do
          begin
            for i := 0 to TFCube(ecran).lDispo.Items.Count-1 do
            begin
              if Pos(TabHisto[j],TFCube(ecran).lDispo.Items.Strings[i]) > 0 then okaxe := okaxe + 1 ;
            end;

            for i := 0 to TFCube(ecran).lLigne.Items.Count-1 do
            begin
              if Pos(TabHisto[j],TFCube(ecran).lLigne.Items.Strings[i]) > 0 then okaxe := okaxe + 1 ;
            end;

            for i := 0 to TFCube(ecran).lValeurs.Items.Count-1 do
            begin
              if Pos(TabHisto[j],TFCube(ecran).lValeurs.Items.Strings[i]) > 0  then okaxe := okaxe + 1 ;
            end;

          end;
        end;
  end; //CASE
  TFCube(ecran).FChoix.ItemIndex :=  IndexFChoix;
  Case ContextBP of
    0,1,2 : if (sTypeDA = 'TYPEDA') then
          begin
            if okaxe = 4 then result := true else result := false;
          end else if okaxe = 3 then result := true else result := false;
    3 : if okaxe = NbrValAff(codeSession) then result := true else result := false;
  end;
end;

{ EVI / Procedure pour modifier la TOB afin de changer les valeurs Libelle
  et FormuleSQL des champs YYY1, YYY2, YYY3 }
procedure TOF_QUFCBPBUDGETPREV.ChangePres(Libelle,FormuleSQL1, FormuleSQL2, FormuleSQL3 : string);
var TobPres, TobTemp: tob;
i,ValAffiche: integer;
begin
  ValAffiche:=1;

  Case ContextBP of
    0,1 : begin //Mode-GC
          if GetControlText('ValAffiche') = 'PTC' then ValAffiche:=1;
          if GetControlText('ValAffiche') = 'QTE' then ValAffiche:=0;
          if GetControlText('ValAffiche') = 'PHT' then ValAffiche:=2;
          if GetControlText('ValAffiche') = 'UHT' then ValAffiche:=3;
          if GetControlText('ValAffiche') = 'UTC' then ValAffiche:=4;
          if GetControlText('ValAffiche') = 'PAH' then ValAffiche:=5;
          if GetControlText('ValAffiche') = 'MAR' then ValAffiche:=6;;
        end;
    2 : begin //Compta
          if GetControlText('ValAffiche') = 'DC1' then ValAffiche:=1;
          if GetControlText('ValAffiche') = 'CD1' then ValAffiche:=2;
        end;
    3 : begin //Paie
          ValAffiche:=0;
        end;
  end;

  // Récupération de la TOB qui contient la présentation
  TobPres := TFCube(ecran).ListePresentation.Current;

  Case ContextBP of
    0,1,2 : begin
            if TobPres <> nil then
            begin
              For i := 0 to (TobPres.Detail.Count-1) do
              begin
                TobTemp := TobPres.Detail.Items[i];
                if Pos('YYY',TobTemp.GetValue('DH_NOMCHAMP')) = 0  then ModifTob(TobTemp,ValAffiche);
              end;

            TFCube(Ecran).delFormule('YYY1');
            //Tob temporaire pour chaque TOB fille (YYY1, 2, 3)
            TobTemp := TobPres.FindFirst(['DH_NOMCHAMP'],['YYY1'], TRUE);
            if TobTemp <> nil then
            begin
              TobTemp.PutValue('TITRE',TraduireMemoire('Delta Réalisé - Prévu'));
              TobTemp.PutValue('FORMULESQL',TabRealise[ValAffiche]+'-'+TabPrevu[ValAffiche]);
            end;
            TFCube(Ecran).addFormule(TraduireMemoire('Delta Réalisé - Prévu'),TabRealise[ValAffiche]+'-'+TabPrevu[ValAffiche],'YYY1');

            TFCube(Ecran).delFormule('YYY2');
            TobTemp := TobPres.FindFirst(['DH_NOMCHAMP'],['YYY2'], TRUE);
            if TobTemp <> nil then
            begin
              TobTemp.PutValue('TITRE',TraduireMemoire('Delta Réalisé - Historique'));
              TobTemp.PutValue('FORMULESQL',FormuleSQL3+'-'+FormuleSQL1);
            end;
            TFCube(Ecran).addFormule(TraduireMemoire('Delta Réalisé - Historique'),TabRealise[ValAffiche]+'-'+TabHisto[ValAffiche],'YYY2');

            if (sTypeDA = 'TYPEDA') then
            begin
              TFCube(Ecran).delFormule('YYY3');
              TobTemp := TobPres.FindFirst(['DH_NOMCHAMP'],['YYY3'], TRUE);
              if TobTemp <> nil then
              begin
                TobTemp.PutValue('TITRE','Delta Total Réalisé - ' + TabLibelle[2]+' Prévu');
                TobTemp.PutValue('FORMULESQL',TabRealise[2]+'+QBQ_CAPREAL2-'+TabPrevu[2]);
              end;
              TFCube(Ecran).addFormule('Delta Total Réalisé - ' + TabLibelle[2]+' Prévu',TabRealise[2]+'+QBQ_CAPREAL2-'+TabPrevu[2],'YYY3')
            end;

            //On vérifie que la présentation est compatile et contient les champs nécessaires, sinon initialise
            if VerifPres(Valaffiche) = false then InitChamp(ValAffiche);

          end
          else InitChamp(ValAffiche);
       end;

    3 : begin
          if TobPres <> nil then
          begin
            For i := 0 to (TobPres.Detail.Count-1) do
            begin
              TobTemp := TobPres.Detail.Items[i];
              if Pos('YYY',TobTemp.GetValue('DH_NOMCHAMP')) = 0  then ModifTob(TobTemp,ValAffiche);
            end;

            TFCube(Ecran).delFormule('YYY1');
            TFCube(Ecran).delFormule('YYY2');
            TFCube(Ecran).delFormule('YYY3');
            TFCube(Ecran).delFormule('Delta Réalisé - Prévu');
            TFCube(Ecran).delFormule('Delta Réalisé - Historique');

            //On vérifie que la présentation est compatbile et contient les champs nécessaires, sinon initialise
            if VerifPres(Valaffiche) = false then InitChamp(ValAffiche);

          end
          else InitChamp(ValAffiche);
       end;

  end;  //CASE

  //Pour finir, on réapplique le filtre sélectionné
  if TobPres <> nil then TFCube(ecran).ForceSelect('',TFCube(ecran).FDimensions.Value)
  else
  //Pour rafraichir la selection alors que la présentation n'est pas définie
  //on fait un ForceSelect sur une présentation qui n'existe pas
  TFCube(ecran).ForceSelect('','PRESENTATION NON DEFINIE PAR DEFAUT') ;

end;

{ EVI / Evenement OnChange de la ComboBox "Valeur affichée" }
procedure TOF_QUFCBPBUDGETPREV.ValAfficheOnChange(Sender: TObject);
begin
  Case ContextBP of
    0,1 : begin //Mode-GC
          if GetControlText('ValAffiche') = 'PTC'
          then ChangePres('CA TTC net','QBQ_CAHISTO', 'QBQ_CAPREVU', 'QBQ_CAREALISE');
          if GetControlText('ValAffiche') = 'QTE'
          then ChangePres('Quantité','QBQ_HISTO', 'QBQ_PREVU', 'QBQ_REALISE');
          if GetControlText('ValAffiche') = 'PHT'
          then ChangePres('CA HT net','QBQ_CAHISTO2', 'QBQ_CAPREVU2', 'QBQ_CAREALISE2');
          if GetControlText('ValAffiche') = 'UHT'
          then ChangePres('CA HT avant remise','QBQ_CAHISTO3', 'QBQ_CAPREVU3', 'QBQ_CAREALISE3');
          if GetControlText('ValAffiche') = 'UTC'
          then ChangePres('CA TTC avant remise','QBQ_CAHISTO4', 'QBQ_CAPREVU4', 'QBQ_CAREALISE4');
          if GetControlText('ValAffiche') = 'PAH'
          then ChangePres('Cumul achat en PAMP','QBQ_CAHISTO5', 'QBQ_CAPREVU5', 'QBQ_CAREALISE5');
          if GetControlText('ValAffiche') = 'MAR'
          then ChangePres('Marge','QBQ_CAHISTO6', 'QBQ_CAPREVU6', 'QBQ_CAREALISE6');
        end;
    2 : begin //Compta
          if GetControlText('ValAffiche') = 'DC1'
          then ChangePres('Débit - Crédit','QBQ_CAHISTO', 'QBQ_CAPREVU', 'QBQ_CAREALISE');
          if GetControlText('ValAffiche') = 'CD1'
          then ChangePres('Crédit - Débit','QBQ_CAHISTO2', 'QBQ_CAPREVU2', 'QBQ_CAREALISE2');
        end;
    3 : begin //Paie
          ChangePres('Paie','QBQ_HISTO', 'QBQ_PREVU', 'QBQ_REALISE');
        end;
  end;
end;

procedure TOF_QUFCBPBUDGETPREV.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_QUFCBPBUDGETPREV.OnDisplay () ;
begin
  Inherited ;

end ;

procedure TOF_QUFCBPBUDGETPREV.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_QUFCBPBUDGETPREV.MyOnAfterFormShow;
begin
 TFCube(ecran).ForceSelect('',TFCube(ecran).FDimensions.Value);
end;

{ EVI / Trap de la touche ESC : éviter de sortir pendant la génération }
procedure TOF_QUFCBPBUDGETPREV.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = VK_ESCAPE) and (InProgress=true) then Key := 0;
  if( (Key = VK_F9) and Assigned(BCherche) ) then BCherche.Click;
end;

procedure TOF_QUFCBPBUDGETPREV.DetailAuJour(VertFields, HorzFields, CellValue: hString);
var dateDelai,ValAxeNiv,LibValAxeNiv,CodeSQLAxe,ValAff,Axes,LibAxes,ReqSQL,WhereSql : hString;
    i, nivMax : integer;
    ConvertError:boolean;
    TabLibAxe : array of string;
    Q : TQuery;

    //Modification AGL 7.0.20.14 : ajout de parenthèses dans les champs
    function TrimComa(Field : hString):hstring;
    begin
      Field := AnsiDequotedStr(Field,'"');
      if ( (IsDelimiter('(',Field,1)) AND (IsDelimiter(')',Field,Length(Field))) ) then
      result := copy(Field,2,length(Field)-2) else result := Field;
    end;

begin
  if DonneMethodeSession(codeSession) = '1' then
  begin

    ConvertError:=false;
    nivMax:=ChercheNivMax(codeSession);
    //Valeur affichée dans le cube
    ValAff:=THValComboBox(GetControl('ValAffiche')).Value;
    dateDelai := trouveArgument(VertFields,'QBQ_DATECT','');
    if dateDelai = '' then dateDelai := trouveArgument(HorzFields,'QBQ_DATECT','');
    dateDelai := StringReplace(DateDelai,'"','',[ rfReplaceAll ]);

    CodeSQLAxe := '';
    Axes:='';

    { EVI / ToDo : récupération des critères restrictifs
    --> GetControlText('QBQ_VALAXECT'+intToStr(i))
    CodeSQLAXE := codeFiltre; }

    //Récupération de la mise en forme
    for i:=1 to nivMax do
    begin
      ValAxeNiv := ''; LibValAxeNiv := '';
      ValAxeNiv := trouveArgument(VertFields,'AXE'+IntToStr(i),'');
      if ValAxeNiv = '' then ValAxeNiv := trouveArgument(HorzFields,'AXE'+IntToStr(i),'');
      if ValAxeNiv = '' then
      begin
        LibValAxeNiv := trouveArgument(VertFields,'LIBAXE'+IntToStr(i),'');
        if LibValAxeNiv = '' then LibValAxeNiv := trouveArgument(HorzFields,'LIBAXE'+IntToStr(i),'');
        if LibValAxeNiv <> '' then
        begin
          CodeSQLAxe := CodeSQLAxe + ' AND QBQ_LIBVALAXECT'+IntToStr(i)+'="'+TrimComa(LibValaxeNiv)+'"';
          LibAxes := LibAxes + 'VALAXENIV'+IntToStr(i)+'='+TrimComa(LibValAxeNiv)+';';
        end;
      end
      else
      begin
        CodeSQLAxe := CodeSQLAxe + ' AND QBQ_VALAXECT'+IntToStr(i)+'="'+TrimComa(ValAxeNiv)+'"';
        Axes := Axes + 'VALAXENIV'+IntToStr(i)+'='+TrimComa(ValAxeNiv)+';';
      end;
      if ReqSQL = '' then  ReqSQL := 'QBQ_VALAXECT'+IntToStr(i)+',QBQ_LIBVALAXECT'+IntToStr(i)
      else ReqSQL := ReqSQL + ',QBQ_VALAXECT'+IntToStr(i)+',QBQ_LIBVALAXECT'+IntToStr(i);
    end;

    if LibAxes <> '' then
    begin
      SetLength(TabLibAxe,nivMax*2);
      WhereSql := StringReplace(ReqSQL,',',' IS NOT NULL AND ', [rfReplaceAll, rfIgnoreCase]) + ' IS NOT NULL';
      Q:=OpenSQL('SELECT ##TOP 1## '+ReqSQL+' FROM QBPCUBETMP WHERE '+WhereSql+' AND QBQ_UTILISATEUR="'+V_PGI.User+'"',true);
      for i := 1 to nivMax*2 do
      begin
        TabLibAxe[i-1] :=  Q.Fields[i-1].AsString;
      end;
      Ferme(Q);
      for i := 1 to nivMax*2 do
      begin
        if Odd(i) then LibAxes := StringReplace(LibAxes,TabLibAxe[i],TabLibAxe[i-1], [rfReplaceAll, rfIgnoreCase]);
      end;
    end;

    Axes := Axes + LibAxes;


    //Champ date vide = Ligne Total
    if dateDelai = ''  then PGIInfo('Aucune date sélectionnée.')
    else
    begin
      dateDelai := VerifDateCube(TrimComa(dateDelai));
      if dateDelai = '' then
      begin
        PGIInfo('Le format de date est incorrect.#13#10 Merci d''utiliser un format jj/mm/aaaa');
        ConvertError:=true;
      end;
    if ConvertError=false then aglLanceFiche('Q', 'QUFVBPDETAILJOURM', '', '','SESSION='+codeSession+';'+Axes+'VALAFF='+ValAff+';DELAI='+dateDelai+';VALEUR='+CellValue);
    end;                                                        
  end;
end;

procedure TOF_QUFCBPBUDGETPREV.BChercheOnClick(Sender: TObject);
begin
  codesession := GetControlText('QBQ_CODESESSION');
  if (ControleSession(codesession)=false)
  then PGIINFO('Le code session doit être renseigné correctement.')
  else TFCube(ecran).BChercheClick(nil);
end;

procedure TOF_QUFCBPBUDGETPREV.VerifValeurs(Sender: TObject);
begin
  if TFCube(Ecran).lValeurs.Count > 7 then
  begin
    PGIINFO('Le quota maximum de 7 valeurs est atteint.') ;
    TFCube(Ecran).lValeurs.Selected[TFCube(Ecran).lValeurs.count-1] := true;
    TFCube(Ecran).lValeurs.OnDblClick(self);
  end;
end;

Initialization
  registerclasses ( [ TOF_QUFCBPBUDGETPREV ] ) ; 
end.
