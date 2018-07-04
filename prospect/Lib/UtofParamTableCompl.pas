unit UtofParamTableCompl;

interface
uses  StdCtrls,Controls,Classes,sysutils,HTB97, Spin
      ,HCtrls,HEnt1,UTOF,UTOB,graphics
      ,grids,windows
{$IFNDEF EAGLCLIENT}
      ,db {$IFNDEF DBXPRESS},dbtables{BDE}{$ELSE},uDbxDataSet{$ENDIF}
{$ENDIF}
      ;

Const
   NbTotalChamps : Integer = 60;
   CleChoixCode : String = 'RTC';
   ChampsNonTraitesRPC : String = 'RPC_RANG;RPC_AUXILIAIRE;RPC_DATECREATION;RPC_DATEMODIF;RPC_CREATEUR;RPC_UTILISATEUR;';
Type
     TOF_PARAMTABLECOMPL = Class (TOF)
     private                 // NbTotalChamps
                         // 1:libellé, 2:type champ, 3:code tablette       
        InfosChampListe: array[1..60,1..4] of string;
        InfosChampEnleves: array[1..60,1..4] of string;
        Oblig : array[1..60] of String;
        LesColonnes : string ;
        BGAUCHE : TToolbarButton97;
        BDROITE : TToolbarButton97;
        BCENTRE : TToolbarButton97;
        BSUPCOL : TToolbarButton97;
        NBDECIMALES : TSpinEdit;
        GS : THGRID ;
        Titres : HTStrings;
        TTitre : THEdit;
        ListeChamps : THListBox;
        TOblig : TCheckBox;
        NbChpsEnleves : Integer;
        procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSDblClick(Sender: TObject);
        procedure GSExchangeCol(Sender: TObject;FromIndex, ToIndex: Longint);
        procedure ListeDblClick(Sender: TObject);
        procedure TitreExit(Sender: TObject);

        procedure DessineCell (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
        procedure ChargeListBox ();
        procedure BGAUCHEClick(Sender: TObject);
        procedure BDROITEClick(Sender: TObject);
        procedure BCENTREClick(Sender: TObject);
        procedure BSUPCOLClick(Sender: TObject);
        procedure NBDECIMALESExit(Sender: TObject);
        procedure MajTitre;
        procedure ObligExit (Sender: TObject);
        Function  ChercherCode ( Code : String ) : String;
        procedure EchangeColonne(FromIndex, ToIndex: Longint);
        procedure AfficheAttributs;
     public
         Action   : TActionFiche ;
         procedure OnArgument (Arguments : String ) ; override ;
         procedure Onclose  ; override ;
         procedure OnLoad ; override ;
         procedure OnUpdate ; override ;
         Procedure OnNew  ; override ;
         Procedure OnDelete ; override ;
     END ;

const colRang=1 ;
      colAuxi=2 ;

implementation

Procedure TOF_PARAMTABLECOMPL.OnArgument (Arguments : String ) ;
var NbCol : integer ;
    St : string ;
begin
inherited ;

LesColonnes:='FIXED' ;
NbCol:=1;

GS:=THGRID(GetControl('G'));
GS.OnCellEnter:=GSCellEnter ;
GS.OnCellExit:=GSCellExit ;
GS.OnRowEnter:=GSRowEnter ;
GS.OnDblClick:=GSDblClick;
GS.OnColumnMoved:=GSExchangeCol;
//GS.GetCellCanvas:= DessineCell;
GS.PostDrawCell:= DessineCell;
GS.ColCount:=NbCol;
//GS.ColAligns[colRang]:=taCenter;
//GS.ColTypes[colRang]:='I' ;
GS.ColWidths[0]:=15;

St:=LesColonnes ;

Titres:=HTStringList.Create;
Titres.Add(' ');
GS.Titres:=Titres;

//AffecteGrid(GS,Action) ;
//TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS) ;

ListeChamps:=THListBox(GetControl('LISTECHAMPS'));
ListeChamps.OnDblClick:=ListeDblClick;
TTitre:=THEdit(GetControl('TITRE'));
TTitre.OnExit:=TitreExit;
BGAUCHE:=TToolbarButton97(GetControl('BGAUCHE'));
BGAUCHE.OnClick:=BGAUCHEClick;
BDROITE:=TToolbarButton97(GetControl('BDROITE'));
BDROITE.OnClick:=BDROITEClick;
BCENTRE:=TToolbarButton97(GetControl('BCENTRE'));
BCENTRE.OnClick:=BCENTREClick;
BSUPCOL:=TToolbarButton97(GetControl('BSUPCOL'));
BSUPCOL.OnClick:=BSUPCOLClick;
NBDECIMALES:=TSpinEdit(GetControl('NBDECIMALES'));
NBDECIMALES.OnExit:=NBDECIMALESExit;
TOblig:=TCheckBox(GetControl('OBLIGATOIRE'));
TOblig.OnExit:=ObligExit;

//SetControlText ('LIBELLETRI','');
//SetControlText ('VALCADRAGE','');

ChargeListBox ();

end;

Procedure TOF_PARAMTABLECOMPL.OnLoad  ;
begin
inherited ;
end;

Procedure TOF_PARAMTABLECOMPL.OnUpdate  ;
var  i : integer ;
     buff,tri,cadre,(*ValChamp,*)StOrdre : String;
     //Q : TQuery;
begin
inherited ;
    if GS.Col <> 0 then
       Oblig[GS.Col]:=GetControlText('OBLIGATOIRE');
    //Q := OpenSQL ('SELECT * FROM CHOIXCOD WHERE CC_TYPE = "'+CleChoixCode+'"',True) ;
    //while Not Q.EOF do
    //begin
      For i := 1 to GS.ColCount-1 do
      begin
        //if InfosChampEnleves[i,3] = Q.FindField('CC_CODE').asstring then
        //begin
             // formatage de la zone cc_libre :
             // 1 : largeur colonne, 2: colonne de tri O/N
             // 3: saisie obligatoire O/N, 4: format de saisie, 5: cadrage
             if GS.SortedCol= i then tri:='X'
             else tri:='-';
             if GS.ColAligns[i] = TaCenter then cadre:='C'
             else if GS.ColAligns[i] = taRightJustify then cadre:='D'
                  else cadre:='G';
             // format= 'B' pour booléen
             if InfosChampEnleves[i,2] = 'BOOLEAN' then GS.ColFormats[i]:='B';
             buff:=IntToStr(GS.ColWidths[i])+';'+tri+';'+Oblig[i]+';'+
                   GS.ColFormats[i]+';'+cadre;
             StOrdre := Format('%.2u', [i]);
             ExecuteSQL( 'UPDATE CHOIXCOD SET CC_LIBELLE = "'+CheckdblQuote(Titres[i])+
             '", CC_LIBRE = "'+buff+'", CC_ABREGE = "'+StOrdre+
             '" WHERE CC_TYPE = "'+CleChoixCode+'" and CC_CODE = "'+InfosChampEnleves[i,3]+ '"') ;
        //end;
      end;
      //if ((i = GS.ColCount) or (i = 0 )) and (Q.FindField('CC_LIBRE').asstring <> '' ) then
        //  begin
          //ValChamp:='';
          For i := 1 to ( NbTotalChamps - GS.ColCount + 1 ) do
              //if InfosChampListe[i,3] = Q.FindField('CC_CODE').asstring then
              //   ValChamp:=InfosChampListe[i,4];

         ExecuteSQL( 'UPDATE CHOIXCOD SET CC_LIBRE = "",CC_ABREGE = "",CC_LIBELLE = "'+InfosChampListe[i,4]+
         '" WHERE CC_TYPE = "'+CleChoixCode+'" and CC_CODE = "'+InfosChampListe[i,3]+ '"') ;
          //end;
      //Q.Next;
    //end;
    //Ferme(Q);
    // libellés tables libres : on mer .- devant le libellé si ce n'est pas déjà fait
    ExecuteSQL( 'UPDATE CHOIXCOD SET CC_LIBELLE = ".-"+CC_LIBELLE '+
    'WHERE CC_TYPE = "'+CleChoixCode+'" and CC_CODE like "TL%" and CC_LIBRE = "" and CC_ABREGE = ""'+
    ' and cc_libelle not like ".-%"') ;

    AvertirTable ('RTLIBTABLECOMPL') ;
end;

Procedure TOF_PARAMTABLECOMPL.OnClose ;
begin
inherited ;
Titres.Free;
//PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB,  0) ;
//application.processMessages;
LastError:=0;
{if GS.rowcount > 2 then LaTob.GetGridDetail( THGRID(GetControl('G')),GS.rowcount-2,'PROSPECTCOMPL',LesColonnes) ;
for i:=0 to LaTob.detail.count-1 do
    begin
    if (LaTob.Detail[i].IsOneModifie) or (TobToDelete.detail.count > 0) then
   //if LaTob.IsOneModifie then
        begin
        if PGIAsk('Voulez-vous enregistrer les modifications ?',Ecran.Caption)=mrYes then
            begin
            LastError:=0;
            OnUpdate;
            if LastError<>0 then exit;
            break;
            end else
            begin
            LastError:=0;
            break;
            end;
        end;
    end;
    
LaTob.free ; Latob:=nil;
TobToDelete.free ;  }
end;


Procedure TOF_PARAMTABLECOMPL.OnNew  ;
begin
inherited ;
GS.InsertRow(GS.row) ;
GS.row:=gs.row-1 ;
if Copy(GS.ColFormats[GS.Col],1,3)='CB=' then
       GS.ShowCombo(GS.Col,GS.Row) ;
// GS.row:=GS.RowCount-1 ;
end;

Procedure TOF_PARAMTABLECOMPL.OnDelete  ;
begin
inherited ;

end;

procedure TOF_PARAMTABLECOMPL.ListeDblClick(Sender: TObject);
var i,j : integer;
begin
    GS.ColCount:=GS.ColCount+1;
    Titres.Add(ListeChamps.Items.Strings[ListeChamps.ItemIndex]);
    MajTitre;

    SetControlText ('TITRE', Titres[GS.ColCount-1]);
    if InfosChampListe[(ListeChamps.ItemIndex)+1,2] = 'DATE' then
        begin
        GS.ColTypes[GS.ColCount-1] := 'D';
        GS.ColFormats[GS.ColCount-1]:=ShortDateFormat;
        end;
    if InfosChampListe[ListeChamps.ItemIndex+1,2] = 'BOOLEAN' then
       GS.ColTypes[GS.ColCount-1] := 'B';
    if InfosChampListe[ListeChamps.ItemIndex+1,2] = 'DOUBLE' then
       begin
       GS.ColTypes[GS.ColCount-1] := 'F';
       GS.ColFormats[GS.ColCount-1]:='#,##0.00';
       end;
    GS.ColAligns[GS.ColCount-1]:=taCenter;
    //SetControlText ('VALCADRAGE','Cadrage centré');
    Inc(NbChpsEnleves);
    for j:=1 to 4 do
        InfosChampEnleves[NbChpsEnleves,j]:=InfosChampListe[ListeChamps.ItemIndex+1,j] ;
    for i:=(ListeChamps.ItemIndex+1 ) to (NbTotalChamps-1) do
        begin
        for j:=1 to 4 do
            InfosChampListe[i,j]:=InfosChampListe[i+1,j];
        Oblig[i]:=Oblig[i+1];
        end;
    for j:=1 to 4 do
        InfosChampListe[(NbTotalChamps-GS.ColCount+2),j]:='';
    ListeChamps.Items.Delete(ListeChamps.ItemIndex);
    GS.Col:=GS.ColCount-1;
end;

procedure TOF_PARAMTABLECOMPL.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
AfficheAttributs;
end;

procedure TOF_PARAMTABLECOMPL.AfficheAttributs;
var i,INum : integer;
    formatch,StLib,Code : string;
begin
// Affectation du titre de la colonne dans le champ 'TITRE'
if GS.Col <> 0 then
   begin
   // no colonne
   SetControlText ('NOCOL',IntToStr(GS.Col));
   // TITRE
   SetControlText ('TITRE', Titres[GS.Col]);
   // CADRAGE
   if GS.ColAligns[GS.Col]=taLeftJustify then
         begin
         //SetControlText ('VALCADRAGE','Cadrage à gauche')
         BGAUCHE.Color:=clSilver;
         BGAUCHE.Flat:=True;
          BCENTRE.Color:=clBtnFace;
          BCENTRE.Flat:=False;
          BDROITE.Color:=clBtnFace;
          BDROITE.Flat:=False;
         end
   else if GS.ColAligns[GS.Col]=taRightJustify then
            begin
            BDROITE.Color:=clSilver;
            BDROITE.Flat:=True;
            BCENTRE.Color:=clBtnFace;
            BCENTRE.Flat:=False;
            BGAUCHE.Color:=clBtnFace;
            BGAUCHE.Flat:=False;
           //SetControlText ('VALCADRAGE','Cadrage à droite')
            end
        else
            begin
            //SetControlText ('VALCADRAGE','Cadrage centré')
            BCENTRE.Color:=clSilver;
            BCENTRE.Flat:=True;
            BDROITE.Color:=clBtnFace;
            BDROITE.Flat:=False;
            BGAUCHE.Color:=clBtnFace;
            BGAUCHE.Flat:=False;
            end
   ;
   // NBDECIMALES
   SetControlText ('NBDECIMALES',IntToStr(0));
   formatch:=GS.ColFormats[GS.Col];
   if formatch <> '' then
     if formatch[1]='#' then
     begin
     i:=Pos('.',formatch);
     //'#,##0.00'
     SetControlText ('NBDECIMALES',IntToStr( Length(formatch)-i ));
     end;

   // OBLIGATOIRE
   SetControlText ('OBLIGATOIRE',Oblig[GS.Col]);

   // NOMCHAMP
   if InfosChampEnleves[GS.Col,2] = 'COMBO' then StLib:='( '+TraduireMemoire('Table libre n°')+ ' ';
   if InfosChampEnleves[GS.Col,2] = 'DATE' then StLib:='( '+TraduireMemoire('Date libre n°')+' ';
   if InfosChampEnleves[GS.Col,2] = 'BOOLEAN' then StLib:='( '+TraduireMemoire('Booléen libre n°')+' ';
   if InfosChampEnleves[GS.Col,2] = 'DOUBLE' then StLib:='( '+TraduireMemoire('Valeur libre n°')+' ';
   if InfosChampEnleves[GS.Col,2] = 'VARCHAR(35)' then StLib:='( '+TraduireMemoire('Champ libre n°')+' ';
   Code:= InfosChampEnleves[GS.Col,3];
   INum:=Ord(Code[3]);
   if INum < 65 then
     INum:=StrToInt(Code[3])
   else
     INum:=Ord(Code[3])-55;
   StLib:=StLib+IntToStr(INum+1)+' )';
   SetControlText ('NOMCHAMP',StLib);

   end;
//if GS.SortedCol > 0 then
//   SetControlText ('LIBELLETRI',Titres[GS.SortedCol]);

end;


procedure TOF_PARAMTABLECOMPL.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
// Affectation du champ 'TITRE' dans le titre de la colonne
{if ACol <> 0 then
  begin
  Titres[ACol] :=  GetControlText ('TITRE');
  MajTitre;
  end;
}
end;


procedure TOF_PARAMTABLECOMPL.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GS.InvalidateRow(ou) ;
end;

procedure TOF_PARAMTABLECOMPL.GSDblClick (Sender: TObject);
begin
if GS.ColTypes[GS.col]='B' then if GS.cells[GS.col,Gs.Row]='X' then GS.cells[GS.col,Gs.Row]:='-' else GS.cells[GS.col,Gs.Row]:='X';

end;

procedure TOF_PARAMTABLECOMPL.GSExchangeCol(Sender: TObject; FromIndex, ToIndex: Longint);
begin
EchangeColonne (FromIndex, ToIndex);
end;

procedure TOF_PARAMTABLECOMPL.EchangeColonne(FromIndex, ToIndex: Longint);
var Tit,SaveOblig : String;
    i,j : integer;
    SaveInfos : array[1..4] of String;
begin
     Tit:=Titres[FromIndex];
     SaveOblig:=Oblig[FromIndex];
     for j:=1 to 4 do
         SaveInfos[j]:= InfosChampEnleves[FromIndex,j];

     if FromIndex < ToIndex then
        begin
        for i:= FromIndex to (ToIndex-1) do
            begin
            Titres[i]:=Titres[i+1];
            Oblig[i]:=Oblig[i+1];
            for j:=1 to 4 do
                InfosChampEnleves[i,j]:=InfosChampEnleves[i+1,j] ;
            end;
        end
     else
        begin
        for i:= FromIndex downto (ToIndex+1) do
            begin
            Titres[i]:=Titres[i-1];
            Oblig[i]:=Oblig[i-1];
            for j:=1 to 4 do
                InfosChampEnleves[i,j]:=InfosChampEnleves[i-1,j]             
            end;
        end ;
     Titres[ToIndex]:=Tit;
     Oblig[ToIndex]:=SaveOblig;
     for j:=1 to 4 do
         InfosChampEnleves[ToIndex,j]:= SaveInfos[j];
     MajTitre;
end;

procedure TOF_PARAMTABLECOMPL.DessineCell(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
  Arect: Trect ;
begin
If Arow < GS.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=GS.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := GS.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = GS.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=ARect.Right-2 ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;

procedure TOF_PARAMTABLECOMPL.ChargeListBox ();
var TobMulti,FilleTobMulti : TOB;
    Q,QQ : TQuery;
    i,NbCrit,NbChps,Ordre,Colonne,x : integer;
    Critere,Valeurs,Nomchamp : String ;
begin
NbChpsEnleves:=0;
NbChps:=0;
Q := OpenSQL('SELECT DH_NOMCHAMP,DH_LIBELLE,DH_TYPECHAMP FROM DECHAMPS WHERE DH_PREFIXE="RPC" order by DH_NUMCHAMP', true);

TobMulti:=TOB.create ('table multi',NIL,-1);
TobMulti.LoadDetailDB('DECHAMPS','','',Q, false);
Ferme(Q);

for i := 0 to TobMulti.Detail.Count-1 do
    begin
    Valeurs:='';
    FilleTobMulti := TobMulti.Detail[i];
    Nomchamp := FilleTobMulti.GetValue('DH_NOMCHAMP');
    x:= Pos(Nomchamp,ChampsNonTraitesRPC);
    if x <> 0 then continue;
    QQ := OpenSQL('SELECT CC_LIBELLE,CC_LIBRE,CC_ABREGE FROM CHOIXCOD WHERE CC_TYPE="'+CleChoixCode+'" and CC_CODE="'+ChercherCode(FilleTobMulti.GetValue('DH_NOMCHAMP'))+'"', true);
    if not QQ.EOF then
      Valeurs := QQ.Fields[1].AsString;
    if Valeurs <> '' then
      begin
      NbCrit:=0;
      Ordre:=StrToInt(QQ.Fields[2].AsString);
      if ( Ordre > ( GS.ColCount-1) ) then
         begin
         repeat
            GS.ColCount:=GS.ColCount+1;
            GS.ColWidths[GS.ColCount-1]:=0;
            Titres.Add(' ');
         until Ordre = ( GS.ColCount - 1 );
         Colonne:=GS.ColCount-1;
         end
      else
      begin
         Colonne:=Ordre;
      end ;
      Titres[Colonne]:=QQ.Fields[0].AsString;
      MajTitre;
      Inc(NbChpsEnleves);
      Repeat
        Inc(NbCrit);
        Critere:=Trim(ReadTokenSt(Valeurs)) ;
        if NbCrit=1 then
           GS.ColWidths[Colonne]:=StrToInt(Critere);
        if NbCrit=2 then
            if (Critere = 'X') then GS.SortedCol:=Colonne;
        if NbCrit=3 then
           if Critere = 'X' then Oblig[Colonne]:='X'
                            else Oblig[Colonne]:='-';
        if NbCrit=4 then
           GS.ColFormats[Colonne]:=Critere;
        if NbCrit=5 then
           begin
           if Critere = 'G' then
              GS.ColAligns[Colonne]:=TaLeftJustify
           else if Critere = 'D' then
                GS.ColAligns[Colonne]:=TaRightJustify
                else  GS.ColAligns[Colonne]:=TaCenter;

          InfosChampEnleves[Colonne,1]:=QQ.Fields[0].AsString;
          InfosChampEnleves[Colonne,2]:=FilleTobMulti.GetValue('DH_TYPECHAMP');
          InfosChampEnleves[Colonne,3]:=ChercherCode(FilleTobMulti.GetValue('DH_NOMCHAMP'));
          InfosChampEnleves[Colonne,4]:=FilleTobMulti.GetValue('DH_LIBELLE');
        end;
      until  NbCrit=5;
    end
    else
    if (FilleTobMulti.GetValue('DH_NOMCHAMP') <> 'RPC_AUXILIAIRE') and
       (FilleTobMulti.GetValue('DH_NOMCHAMP') <> 'RPC_RANG') then
        begin
        Inc(NbChps);
        InfosChampListe[NbChps,1]:=FilleTobMulti.GetValue('DH_LIBELLE');
        InfosChampListe[NbChps,2]:=FilleTobMulti.GetValue('DH_TYPECHAMP');
        InfosChampListe[NbChps,3]:=ChercherCode(FilleTobMulti.GetValue('DH_NOMCHAMP'));
        InfosChampListe[NbChps,4]:=FilleTobMulti.GetValue('DH_LIBELLE');

        ListeChamps.Items.Add(FilleTobMulti.GetValue('DH_LIBELLE'));
        //Oblig[NbChps]:='X';
        end;
    Ferme(QQ);
    end;
TobMulti.Free;
end;

procedure TOF_PARAMTABLECOMPL.TitreExit(Sender: TObject);
begin
if GS.Col <> 0 then
  begin
  Titres[GS.Col] :=  GetControlText ('TITRE');
  MajTitre;
  end;
end;

procedure TOF_PARAMTABLECOMPL.BGAUCHEClick(Sender: TObject);
begin
GS.ColAligns[GS.Col]:=taLeftJustify;
BGAUCHE.Color:=clSilver;
BGAUCHE.Flat:=True;
BCENTRE.Color:=clBtnFace;
BCENTRE.Flat:=False;
BDROITE.Color:=clBtnFace;
BDROITE.Flat:=False;
//SetControlText ('VALCADRAGE','Cadrage à gauche');
end;

procedure TOF_PARAMTABLECOMPL.BDROITEClick(Sender: TObject);
begin
GS.ColAligns[GS.Col]:=taRightJustify;
BGAUCHE.Color:=clBtnFace;
BGAUCHE.Flat:=False;
BCENTRE.Color:=clBtnFace;
BCENTRE.Flat:=False;
BDROITE.Color:=clSilver;
BDROITE.Flat:=True;
//SetControlText ('VALCADRAGE','Cadrage à droite');
end;

procedure TOF_PARAMTABLECOMPL.BCENTREClick(Sender: TObject);
begin
GS.ColAligns[GS.Col]:=taCenter;
BGAUCHE.Color:=clBtnFace;
BCENTRE.Color:=clSilver;
BDROITE.Color:=clBtnFace;
BGAUCHE.Flat:=False;
BCENTRE.Flat:=True;
BDROITE.Flat:=False;
//SetControlText ('VALCADRAGE','Cadrage centré');
end;

procedure TOF_PARAMTABLECOMPL.BSUPCOLClick(Sender: TObject);
var j : integer;
begin
if GS.Col <> 0 then
   begin
   // MAJ ListeBox
   ListeChamps.Items.Add(InfosChampEnleves[GS.Col,4]);
   // MAJ Infos Champs ListBox
   for j:=1 to 4 do
      InfosChampListe[(ListeChamps.Items.Count),j]:=InfosChampEnleves[GS.Col,j];
   // libellé paramétré = libellé Initial   
   InfosChampListe[(ListeChamps.Items.Count),1]:=InfosChampListe[(ListeChamps.Items.Count),4];
   // MAJ Infos Champs Affichés
   {for i:= GS.Col to NbChpsEnleves - 1 do
      for j:=1 to 4 do
          InfosChampEnleves[i,j]:=InfosChampEnleves[i+1,j] ;
   // MAJ Titres et Obligatoire
   for i:= GS.Col to (GS.ColCount-2) do
        begin
        Titres[i]:=Titres[i+1];
        Oblig[i]:=Oblig[i+1];
        end;
}

   EchangeColonne( GS.Col,GS.ColCount-1 );
   {MoveColumn(GS.Col,GS.ColCount-1) ;}
   GS.Cols[GS.ColCount-1].Clear ;
   GS.ColCount:=GS.ColCount-1 ;
   for j:=1 to 4 do
       InfosChampEnleves[NbChpsEnleves,j]:='';
   Dec(NbChpsEnleves);
   AfficheAttributs;
   Titres.Delete(GS.ColCount);
   end;
end;

procedure TOF_PARAMTABLECOMPL.NBDECIMALESExit(Sender: TObject);
var formatchp : string;
    i : integer;
begin
formatchp:='#,##0.';
for i:= 1 to StrToInt(GetControlText('NBDECIMALES')) do
    formatchp:=formatchp+'0';
GS.ColFormats[GS.Col]:=formatchp;
end;

procedure TOF_PARAMTABLECOMPL.ObligExit(Sender: TObject);
begin
if GS.Col <> 0 then
   Oblig[GS.Col]:=GetControlText('OBLIGATOIRE');
end;

procedure TOF_PARAMTABLECOMPL.MajTitre ;
var F : array[1..60] of string;
    A : array[1..60] of TAlignment;
    T : array[1..60] of Char;
    i : integer;
begin
for i:=1 to GS.ColCount-1 do begin
F[i]:= GS.ColFormats[i];
A[i]:= GS.ColAligns[i];
T[i]:= GS.ColTypes[i];
end;
GS.Titres:=Titres;
for i:=1 to GS.ColCount-1 do begin
GS.ColFormats[i]:= F[i];
GS.ColAligns[i]:= A[i];
GS.ColTypes[i]:= T[i];
end;
end;

Function TOF_PARAMTABLECOMPL.ChercherCode ( Code : String ) : String;
begin
     if copy(Code,1,15) = 'RPC_RPCLIBTABLE' then
        Result := 'TL'+copy(Code,16,1)
     else
       if copy(Code,1,15) = 'RPC_RPCLIBTEXTE'  then
          Result := 'TX'+copy(Code,16,1)
       else
         if copy(Code,1,13) = 'RPC_RPCLIBVAL' then
            Result := 'VL'+copy(Code,14,1)
         else
           if copy(Code,1,14) = 'RPC_RPCLIBBOOL' then
              Result := 'BL'+copy(Code,15,1)
         else
              Result := 'DL'+copy(Code,15,1)
         ;
end;


Initialization
registerclasses([TOF_PARAMTABLECOMPL]);
end.
