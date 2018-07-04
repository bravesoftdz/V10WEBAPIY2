{***********UNITE*************************************************
Auteur  ...... : CATALA David
Créé le ...... : 09/02/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : DPAGRICOLE (DPAGRICOLE)
Mots clefs ... : TOM;DPAGRICOLE
*****************************************************************}
Unit DPAGRICOLE_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,dbCtrls,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
     fe_main,
{$else}
     eFiche,
     UtileAGL,
     eFichList,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     HTB97,
     UTob;

Type
  TOM_DPAGRICOLE = Class (TOM)
    procedure OnArgument ( S: String ) ; override;
    procedure OnAfterUpdateRecord ; override;
    procedure OnLoadRecord ; override;
    procedure OnClose ; override ;

   private
    BChargement    : Boolean;
    BModification  : Boolean;
    procedure ControlOnClickNiv1 (Sender : TObject);
    procedure ControlOnClickNiv2 (Sender : TObject);
    procedure ControlOnClickCultureAnnuelle (Sender : TObject);
    procedure ControlOnClickCulturePerenne (Sender : TObject);

    function  IsControlNiv1 (NomControle : String) : Boolean;
    function  IsControlNiv2 (NomControle : String) : Boolean;
    procedure InitialiserNiv1 (UnCodeNiv : Integer);
    procedure InitialiserNiv2 (UnCodeNiv : Integer);
    procedure InitialiserDetail (UnCodeNiv : Integer);
  end ;

Implementation

uses UDossierSelect,DpJurOutils;

Const
    NbreElement  : Array [1..5] of Integer =(6,5,13,6,4);
    ListeCbNiv1  : Array [1..5]  of string =('DAG_100CULTURE','DAG_200PERENNE','DAG_300ANIMAUX','DAG_400TRANSF','DAG_500ACTIVAUTRE');

    ListeCbNiv21 : Array [1..6]  of String =('DAG_101CEREALE','DAG_102OLEOPROTE','DAG_103FOURRAGE','DAG_104LEGUME','DAG_105FLORALE','DAG_106AUTRE');
    ListecbNiv22 : Array [1..5]  of String =('DAG_201FRUITIERE','DAG_202VITICULTURE','DAG_203SYLVICULT','DAG_204PEPINIERE','DAG_205AUTRE');
    ListeCbNiv23 : Array [1..13] of String =('DAG_301EQUIN','DAG_302CAPRIN','DAG_303OVIN','DAG_304BOVINS','','','DAG_307PORC','DAG_308AVICULTUREC','DAG_309AVICULTUREO','DAG_310GIBIER','DAG_311COQUILLAGE','DAG_312FERMEAQUA','DAG_313AUTRE');
    ListeCbNiv24 : Array [1..6]  of String =('DAG_401PRODUCTIONL','DAG_402PRODUCTIONB','DAG_403GAVAGE','DAG_404PRODUCTIONB','DAG_405PRODUCTIONF','DAG_406AUTRE');
    ListeCbNiv25 : Array [1..4]  of String =('DAG_501ACTIVITEV','DAG_502STOCKAGE','DAG_503TOURISME','DAG_504PRESTATION');

//----------------------------
//--- Nom : OnArgument
//----------------------------
procedure TOM_DPAGRICOLE.OnArgument ( S: String );
begin
 BChargement:=True;
 TToolBarButton97 (GetControl ('BINSERT')).Visible:=False;
 TToolBarButton97 (GetControl ('BDELETE')).Visible:=False;

 Inherited ;

 //--- Niveau1
 THCheckBox(GetControl(ListeCbNiv1 [1])).OnClick := ControlOnClickNiv1;
 THCheckBox(GetControl(ListeCbNiv1 [2])).OnClick := ControlOnClickNiv1;
 THCheckBox(GetControl(ListeCbNiv1 [3])).OnClick := ControlOnClickNiv1;
 THCheckBox(GetControl(ListeCbNiv1 [4])).OnClick := ControlOnClickNiv1;
 THCheckBox(GetControl(ListeCbNiv1 [5])).OnClick := ControlOnClickNiv1;

 //--- Niveau21
 THCheckBox(GetControl(ListeCbNiv21 [1])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv21 [2])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv21 [3])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv21 [4])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv21 [5])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv21 [6])).OnClick := ControlOnClickNiv2;

 //--- Niveau22
 THCheckBox(GetControl(ListeCbNiv22 [1])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv22 [2])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv22 [3])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv22 [4])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv22 [5])).OnClick := ControlOnClickNiv2;

 //--- Niveau23
 THCheckBox(GetControl(ListeCbNiv23 [1])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv23 [2])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv23 [3])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv23 [4])).OnClick := ControlOnClickNiv2;
 //THCheckBox(GetControl(ListeCbNiv23 [5])).OnClick := ControlOnClickNiv2;
 //THCheckBox(GetControl(ListeCbNiv23 [6])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv23 [7])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv23 [8])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv23 [9])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv23 [10])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv23 [11])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv23 [12])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv23 [13])).OnClick := ControlOnClickNiv2;

 //--- Niveau24
 THCheckBox(GetControl(ListeCbNiv24 [1])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv24 [2])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv24 [3])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv24 [4])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv24 [5])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv24 [6])).OnClick := ControlOnClickNiv2;

 //--- Niveau25
 THCheckBox(GetControl(ListeCbNiv25 [1])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv25 [2])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv25 [3])).OnClick := ControlOnClickNiv2;
 THCheckBox(GetControl(ListeCbNiv25 [4])).OnClick := ControlOnClickNiv2;

 if (GetControl('BCULTUREANNUELLE')<>nil) then TToolBarButton97 (GetControl ('BCULTUREANNUELLE')).Onclick:=ControlOnClickCultureAnnuelle;
 if (GetControl('BCULTUREPERENNE')<>nil) then TToolBarButton97 (GetControl ('BCULTUREPERENNE')).Onclick:=ControlOnClickCulturePerenne;
end ;

//-------------------------
//--- Nom : OnClose
//-------------------------
procedure TOM_DPAGRICOLE.OnClose;
begin
 TFFiche(Ecran).retour := BoolToStr (BModification);
end;

//-------------------------
//--- Nom : OnLoadRecord
//-------------------------
procedure TOM_DPAGRICOLE.OnLoadRecord;
var IndiceNiv1, IndiceNiv2 : Integer;
begin
 BChargement:=True;
 Inherited;

 //--- Initialisation de l'onglet Activité
 for IndiceNiv1:=1 to 5 do
  begin
   InitialiserNiv1 (100*IndiceNiv1);
   for IndiceNiv2:=1 to NbreElement [IndiceNiv1] do
    InitialiserNiv2 ((100*IndiceNiv1)+IndiceNiv2);
  end;

 //--- Activation des boutons d'accés aux parcelles
 if (THCheckBox (GetControl ('DAG_100CULTURE')).Checked) then
  TToolBarButton97 (GetControl('BCULTUREANNUELLE')).Enabled:=True
 else
  TToolBarButton97 (GetControl('BCULTUREANNUELLE')).Enabled:=False;

 if (THCheckBox (GetControl ('DAG_200PERENNE')).Checked) then
  TToolBarButton97 (GetControl('BCULTUREPERENNE')).Enabled:=True
 else
  TToolBarButton97 (GetControl('BCULTUREPERENNE')).Enabled:=False;

 //--- Initialisation de l'onglet Complément
 InitialiserDetail (601);
 InitialiserDetail (602);
 InitialiserDetail (603);
 InitialiserDetail (604);
 InitialiserDetail (605);
 InitialiserDetail (606);

 ModeEdition (DS);
 BChargement   :=False;
 BModification :=False;
end;

//-------------------------------
//--- Nom : ControlOnClickNiv1
//-------------------------------
procedure TOM_DPAGRICOLE.ControlOnClickNiv1 (Sender : TObject);
begin
 if not(BChargement) and (IsControlNiv1 ((Ecran.ActiveControl).Name))then
  InitialiserNiv1 (StrToInt (Copy ((Ecran.ActiveControl).Name,5,3)));
end;

//-------------------------------
//--- Nom : ControlOnClickNiv2
//-------------------------------
procedure TOM_DPAGRICOLE.ControlOnClickNiv2 (Sender : TObject);
begin
 if not(BChargement) and (IsControlNiv2 ((Ecran.ActiveControl).Name)) then
  InitialiserNiv2 (StrToInt (Copy ((Ecran.ActiveControl).Name,5,3)));
end;

//------------------------------------------
//--- Nom : ControlOnClickCultureAnnuelle
//------------------------------------------
procedure TOM_DPAGRICOLE.ControlOnClickCultureAnnuelle (Sender : TObject);
var LancerFicheParcelle : Boolean;
    ListeCode           : String;
    Indice,IndiceNiv3   : Integer;
    SSql                : String;
    UneTob              : Tob;
begin
 LancerFicheParcelle:=False;
 //--- La fiche est en mode consultation
 if (DS.State=dsBrowse) then
  LancerFicheParcelle:=True;

 //--- La fiche n'est pas en mode consultation
 if (DS.State<>dsBrowse) then
  if TFFiche (Ecran).EnregOk then
   begin
    TFFiche(Ecran).Bouge(NbPost);
    LancerFicheParcelle:=True;
   end;

 if (LancerFicheParcelle) then
  begin
   ListeCode:='';
   for Indice:= 1 to 6 do
    begin
     if (THMultiValCombobox (GetControl ('CBDETAIL10'+IntToStr (Indice))).Text<>'') then
      begin
       if (THMultiValCombobox (GetControl ('CBDETAIL10'+IntToStr (Indice))).Text='<<Tous>>') then
        begin
         SSql:='SELECT YDS_CODE FROM CHOIXDPSTD WHERE YDS_PREDEFINI="DOS" AND YDS_NODOSSIER="'+VH_DOSS.NoDossier+'" AND YDS_TYPE="DAG" AND YDS_CODE LIKE "10'+IntToStr (Indice)+'%"';
         UneTob := Tob.Create('Niveau 3', nil, -1);
         UneTob.LoadDetailFromSQL(SSql);

         for IndiceNiv3:=0 to UneTob.Detail.Count-1 do
          ListeCode:=ListeCode+UneTob.Detail[IndiceNiv3].GetValue('YDS_CODE')+';';
         UneTob.Free;
        end
       else
        ListeCode:=ListeCode+THMultiValCombobox (GetControl ('CBDETAIL10'+IntToStr (Indice))).Text;
      end;
    end;

   ListeCode:=StringReplace (ListeCode,';','|',[rfReplaceAll]);
   AGLLanceFiche ('DP','FICHPARCELLE','','',VH_DOSS.NODOSSIER+';-;'+ListeCode);
  end;
end;

//------------------------------------------
//--- Nom : ControlOnClickCulturePerenne
//------------------------------------------
procedure TOM_DPAGRICOLE.ControlOnClickCulturePerenne (Sender : TObject);
var LancerFicheParcelle : Boolean;
    ListeCode           : String;
    Indice,IndiceNiv3   : Integer;
    SSql                : String;
    UneTob              : Tob;
begin
 LancerFicheParcelle:=False;
 //--- La fiche est en mode consultation
 if (DS.State=dsBrowse) then
  LancerFicheParcelle:=True;

 //--- La fiche n'est pas en mode consultation
 if (DS.State<>dsBrowse) then
  if TFFiche (Ecran).EnregOk then
   begin
    TFFiche(Ecran).Bouge(NbPost);
    LancerFicheParcelle:=True;
   end;

 if (LancerFicheParcelle) then
  begin
   ListeCode:='';
   for Indice:= 1 to 5 do
    begin
     if (THMultiValCombobox (GetControl ('CBDETAIL20'+IntToStr (Indice))).Text<>'') then
      begin
       if (THMultiValCombobox (GetControl ('CBDETAIL20'+IntToStr (Indice))).Text='<<Tous>>') then
        begin
         SSql:='SELECT YDS_CODE FROM CHOIXDPSTD WHERE YDS_PREDEFINI="DOS" AND YDS_NODOSSIER="'+VH_DOSS.NoDossier+'" AND YDS_TYPE="DAG" AND YDS_CODE LIKE "20'+IntToStr (Indice)+'%"';
         UneTob := Tob.Create('Niveau 3', nil, -1);
         UneTob.LoadDetailFromSQL(SSql);

         for IndiceNiv3:=0 to UneTob.Detail.Count-1 do
          ListeCode:=ListeCode+UneTob.Detail[IndiceNiv3].GetValue('YDS_CODE')+';';
         UneTob.Free;
        end
       else
        ListeCode:=ListeCode+THMultiValCombobox (GetControl ('CBDETAIL20'+IntToStr (Indice))).Text;
      end;
    end;

   ListeCode:=StringReplace (ListeCode,';','|',[rfReplaceAll]);
   AGLLanceFiche ('DP','FICHPARCELLE','','',VH_DOSS.NODOSSIER+';X;'+ListeCode);
  end;
end;

//-------------------------------
//--- Nom : IsControlNiv1
//-------------------------------
function TOM_DPAGRICOLE.IsControlNiv1 (NomControle : String) : Boolean;
var IndiceNiv1 : Integer;
begin
 Result:=False;
 for IndiceNiv1:=1 to 5 do
  if (ListeCbNiv1 [IndiceNiv1]=NomControle) then
   begin
    Result:=True;
    exit;
   end;
end;

//-------------------------------
//--- Nom : IsControlNiv2
//-------------------------------
function TOM_DPAGRICOLE.IsControlNiv2 (NomControle : String) : Boolean;
var IndiceNiv1,IndiceNiv2 : Integer;
begin
 Result:=False;
 for IndiceNiv1:=1 to 5 do
  begin
   for IndiceNiv2:=1 to NbreElement [IndiceNiv1] do
    begin
     case (IndiceNiv1) of
      1 : if (ListeCbNiv21 [IndiceNiv2]=NomControle) then begin Result:=True;Exit;end;
      2 : if (ListeCbNiv22 [IndiceNiv2]=NomControle) then begin Result:=True;Exit;end;
      3 : if (ListeCbNiv23 [IndiceNiv2]=NomControle) then begin Result:=True;Exit;end;
      4 : if (ListeCbNiv24 [IndiceNiv2]=NomControle) then begin Result:=True;Exit;end;
      5 : if (ListeCbNiv25 [IndiceNiv2]=NomControle) then begin Result:=True;Exit;end;
     end;
    end;
  end;
end;

//---------------------------------------
//--- Nom   : InitialiserNiv1
//---------------------------------------
procedure TOM_DPAGRICOLE.InitialiserNiv1 (UnCodeNiv : Integer);
var NumChp, NomChpNiv1, NomChpNiv2 : String;
    IndiceNiv2                     : Integer;
    Etat                           : Boolean;
begin
 NomChpNiv1:=ListeCbNiv1 [UnCodeNiv div 100];
 Etat:=THCheckBox(GetControl(NomChpNiv1)).Checked;

 if (NomChpNiv1='DAG_100CULTURE') then
  begin
   if not (THCheckBox(GetControl('DAG_200PERENNE')).Checked) then
    THMultiValCombobox (GetControl ('CBDETAIL602')).Enabled:=Etat;

   THLabel (GetControl ('TDAG_CERTIFICATIONC')).Enabled:=Etat;
   TToolBarButton97 (GetControl('BCULTUREANNUELLE')).Enabled:=Etat;
  end;

 if (NomChpNiv1='DAG_200PERENNE') then
  begin
   if not (THCheckBox(GetControl('DAG_100CULTURE')).Checked) then
    THMultiValCombobox (GetControl ('CBDETAIL602')).Enabled:=Etat;

   THLabel (GetControl ('TDAG_CERTIFICATIONC')).Enabled:=Etat;
   TToolBarButton97 (GetControl('BCULTUREPERENNE')).Enabled:=Etat;
  end;

 if (NomChpNiv1='DAG_300ANIMAUX') then
  begin
   THMultiValCombobox (GetControl ('CBDETAIL603')).Enabled:=Etat;
   THLabel (GetControl ('TDAG_CERTIFICATIONA')).Enabled:=Etat;
  end;

 for IndiceNiv2:=1 to NbreElement [UnCodeNiv div 100] do
  begin
   Case (UnCodeNiv div 100) of
    1 : NomChpNiv2:=ListeCbNiv21 [IndiceNiv2];
    2 : NomChpNiv2:=ListeCbNiv22 [IndiceNiv2];
    3 : NomChpNiv2:=ListeCbNiv23 [IndiceNiv2];
    4 : NomChpNiv2:=ListeCbNiv24 [IndiceNiv2];
    5 : NomChpNiv2:=ListeCbNiv25 [IndiceNiv2];
   else NomChpNiv2:='';
   end;

   if (NomChpNiv2<>'') then
    begin
     THCheckBox(GetControl(NomChpNiv2)).Enabled:=Etat;
     if (not Etat) then
      begin
       THCheckBox(GetControl(NomChpNiv2)).Checked:=FALSE;
       SetField (NomChpNiv2,'-');
       NumChp:=Copy (NomChpNiv2,5,3);
       THMultiValCombobox (GetControl ('CBDETAIL'+NumChp)).Enabled:=FALSE;
       THMultiValCombobox (GetControl ('CBDETAIL'+NumChp)).Text:='';
      end;
    end;
  end;
end;

//---------------------------------------
//--- Nom : InitialiserNiv2
//---------------------------------------
procedure TOM_DPAGRICOLE.InitialiserNiv2 (UnCodeNiv : Integer);
var NumChp, NomChpNiv2 : String;
    Etat       : Boolean;
begin
 Case (UnCodeNiv div 100) of
  1 : NomChpNiv2:=ListeCbNiv21 [UnCodeNiv mod 100];
  2 : NomChpNiv2:=ListeCbNiv22 [UnCodeNiv mod 100];
  3 : NomChpNiv2:=ListeCbNiv23 [UnCodeNiv mod 100];
  4 : NomChpNiv2:=ListeCbNiv24 [UnCodeNiv mod 100];
  5 : NomChpNiv2:=ListeCbNiv25 [UnCodeNiv mod 100];
 else NomChpNiv2:='';
 end;

 if (NomChpNiv2<>'') then
  begin
   Etat:=THCheckBox(GetControl(NomChpNiv2)).Checked;
   NumChp:=Copy (NomChpNiv2,5,3);
   THMultiValCombobox (GetControl ('CBDETAIL'+NumChp)).Enabled:=Etat;
   if (Etat) then
    InitialiserDetail (UnCodeNiv)
   else
    THMultiValCombobox (GetControl ('CBDETAIL'+NumChp)).Text:='';
  end;

 if (NomChpNiv2='DAG_104LEGUME') then
  begin
   Etat:=THCheckBox(GetControl(NomChpNiv2)).Checked;
   THMultiValCombobox (GetControl ('CBDETAIL604')).Enabled:=Etat;
   THLabel (GetControl ('TDAG_MODELEGUME')).Enabled:=Etat;
  end;

 if (NomChpNiv2='DAG_304BOVINS') then
  begin
   Etat:=THCheckBox(GetControl(NomChpNiv2)).Checked;
   THMultiValCombobox (GetControl ('CBDETAIL605')).Enabled:=Etat;
   THLabel (GetControl ('TDAG_RACE')).Enabled:=Etat;
  end;

 if (NomChpNiv2='DAG_307PORC') then
  begin
   Etat:=THCheckBox(GetControl(NomChpNiv2)).Checked;
   THMultiValCombobox (GetControl ('CBDETAIL606')).Enabled:=Etat;
   THLabel (GetControl ('TDAG_ACTIVITEPORC')).Enabled:=Etat;
  end;
end;

//---------------------------------------
//--- Nom : InitialiserDetail
//---------------------------------------
procedure TOM_DPAGRICOLE.InitialiserDetail (UnCodeNiv : Integer);
var IndiceNiv3            : Integer;
    SSql,SListeCodeDetail : String;
    UneTob                : Tob;
begin
 SSql:='SELECT YDS_CODE FROM CHOIXDPSTD WHERE YDS_PREDEFINI="DOS" AND YDS_NODOSSIER="'+VH_DOSS.NoDossier+'" AND YDS_TYPE="DAG" AND YDS_CODE LIKE "'+IntToStr (UnCodeNiv)+'%"';
 UneTob := Tob.Create('Niveau 3', nil, -1);
 UneTob.LoadDetailFromSQL(SSql);

 SListeCodeDetail:='';
 for IndiceNiv3:=0 to UneTob.Detail.Count-1 do
  SListeCodeDetail:=SListeCodeDetail+UneTob.Detail[IndiceNiv3].GetValue('YDS_CODE')+';';
 UneTob.Free;

 THMultiValCombobox (GetControl ('CBDETAIL'+IntToStr (UnCodeNiv))).Text:=SListeCodeDetail;
end;

//---------------------------------------
//--- Nom : OnAfterUpdateRecord
//---------------------------------------
procedure TOM_DPAGRICOLE.OnAfterUpdateRecord;
var IndiceNiv1,IndiceNiv2,IndiceNiv3       : Integer;
    NomChpDetail,NomChpNiv2,SListeCodeNiv3 : String;
    CodeNiv3,LibelleNiv3                   : String;
    Indice,Indice1,Indice2                 : Integer;
    Code,Libelle,SListeCode                : String;
begin
 //--- Suppression de tout les enregistrements du dossier courant dans la table CHOIXDPSTD
 ExecuteSQL ('DELETE FROM CHOIXDPSTD WHERE YDS_TYPE="DAG" AND YDS_PREDEFINI="DOS" AND YDS_NODOSSIER="'+VH_DOSS.NoDossier+'"');

 //--- Sauvegarde de tout les MultiValComboBox de l'onglet ACTIVITE dans la table ChoixDPSTD
 for IndiceNiv1:=1 to 5 do
  for IndiceNiv2:=1 to NbreElement [IndiceNiv1] do
   begin
    Case IndiceNiv1 of
     1 : NomChpNiv2:=ListeCbNiv21 [IndiceNiv2];
     2 : NomChpNiv2:=ListeCbNiv22 [IndiceNiv2];
     3 : NomChpNiv2:=ListeCbNiv23 [IndiceNiv2];
     4 : NomChpNiv2:=ListeCbNiv24 [IndiceNiv2];
     5 : NomChpNiv2:=ListeCbNiv25 [IndiceNiv2];
    else NomChpNiv2:='';
    end;

    if (NomChpNiv2<>'') then
     begin
      if (THCheckBox(GetControl(NomChpNiv2)).Checked) then
       begin
        NomChpDetail:='CBDETAIL'+Copy (NomChpNiv2,5,3);

        //--- Tous les éléments sont sélectionnés
        if (THMultiValComboBox (GetControl (NomChpDetail)).Text='') or
           (THMultiValComboBox (GetControl (NomChpDetail)).Text='<<Tous>>') then
         begin
          for IndiceNiv3:=0 to THMultiValCombobox (GetControl (NomChpDetail)).items.count-1 do
           begin
            CodeNiv3:=THMultiValCombobox (GetControl (NomChpDetail)).Values [IndiceNiv3];
            LibelleNiv3:=THMultiValCombobox (GetControl (NomChpDetail)).items [IndiceNiv3];
            ExecuteSQl ('INSERT INTO CHOIXDPSTD (YDS_PREDEFINI,YDS_NODOSSIER,YDS_TYPE,YDS_CODE,YDS_LIBELLE,YDS_ABREGE,YDS_LIBRE) values ("DOS","'+VH_DOSS.NoDossier+'","DAG","'+CodeNiv3+'","'+LibelleNiv3+'","","")');
           end;
         end
        else
         //--- Un ou plusieurs sont sélectionnés
         begin
          SListeCodeNiv3:=THMultiValCombobox (GetControl (NomChpDetail)).Text;
          CodeNiv3:=ReadToKenSt (SListeCodeNiv3);
          while (CodeNiv3<>'') do
           begin
            IndiceNiv3:=THMultiValCombobox (GetControl (NomChpDetail)).Values.IndexOf (CodeNiv3);
            LibelleNiv3:=THMultiValCombobox (GetControl (NomChpDetail)).items [IndiceNiv3];
            ExecuteSQl ('INSERT INTO CHOIXDPSTD (YDS_PREDEFINI,YDS_NODOSSIER,YDS_TYPE,YDS_CODE,YDS_LIBELLE,YDS_ABREGE,YDS_LIBRE) values ("DOS","'+VH_DOSS.NoDossier+'","DAG","'+CodeNiv3+'","'+LibelleNiv3+'","","")');
            CodeNiv3:=ReadToKenSt (SListeCodeNiv3)
           end;
         end;

       end;
     end;
   end;

 //--- Sauvegarde de tout les MultiValComboBox de l'onglet Complement dans la table ChoixDPSTD
 for Indice1:=1 to 6 do
  begin
   NomChpDetail:='CBDETAIL'+IntToStr (600+Indice1);

   if (THMultiValComboBox (GetControl (NomChpDetail)).Enabled) then
    begin
     //--- Tous les éléments sont sélectionnés
     if (THMultiValComboBox (GetControl (NomChpDetail)).Text='') or
        (THMultiValComboBox (GetControl (NomChpDetail)).Text='<<Tous>>') then
      begin
       for Indice2:=0 to THMultiValCombobox (GetControl (NomChpDetail)).items.count-1 do
        begin
         Code:=THMultiValCombobox (GetControl (NomChpDetail)).Values [Indice2];
         Libelle:=THMultiValCombobox (GetControl (NomChpDetail)).items [Indice2];
         ExecuteSQl ('INSERT INTO CHOIXDPSTD (YDS_PREDEFINI,YDS_NODOSSIER,YDS_TYPE,YDS_CODE,YDS_LIBELLE,YDS_ABREGE,YDS_LIBRE) values ("DOS","'+VH_DOSS.NoDossier+'","DAG","'+Code+'","'+Libelle+'","","")');
        end;
      end
     else
      //--- Un ou plusieurs sont sélectionnés
      begin
       SListeCode:=THMultiValCombobox (GetControl (NomChpDetail)).Text;
       Code:=ReadToKenSt (SListeCode);
       while (Code<>'') do
        begin
         Indice:=THMultiValCombobox (GetControl (NomChpDetail)).Values.IndexOf (Code);
         Libelle:=THMultiValCombobox (GetControl (NomChpDetail)).items [Indice];
         ExecuteSQl ('INSERT INTO CHOIXDPSTD (YDS_PREDEFINI,YDS_NODOSSIER,YDS_TYPE,YDS_CODE,YDS_LIBELLE,YDS_ABREGE,YDS_LIBRE) values ("DOS","'+VH_DOSS.NoDossier+'","DAG","'+Code+'","'+Libelle+'","","")');
         Code:=ReadToKenSt (SListeCode);
        end;
      end;
    end;
  end;

 //--- Mise a jour des sections pour la compta
 BModification:=True;
end;



Initialization
  registerclasses ( [ TOM_DPAGRICOLE ] ) ;
end.

