{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 15/04/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFSLARBRE ()
Mots clefs ... : TOF;QUFSLARBRE
*****************************************************************}
Unit QUFSLARBRE_TOF ;

Interface

Uses UTOF,ExtCtrls,ImgList,uTableFiltre,UTob,HEnt1,ComCtrls,HCtrls;

Type
  TOF_QUFSLARBRE = Class (TOF)
    Img0,Img1,Img2,Img3,Img4 : TImage;
    ImgSel : TCustomImageList;
    TF: TTableFiltre;
    TV : THTreeView;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    codeSession,consult,BPInitialise,TabActive,OldCodeAxe:hString;
    nivMax,NbValAff,ValeurAffiche,NbIntervalle,NodeDelete:integer;
    sousNivOk,OkDuplicationRefresh,OkSessionObj,OkSessionInitCoeff,OkSessionInitPrev:boolean;
    OkMaillageJour:boolean;
    TabCodeAxe,TabNumValAxe:array [1..15] of hString;
    TabValAff,TabLibValAff:array [0..10] of hString;
    TabValues : array[1..7] of String;
    TabRefSQLValues : array[1..14] of String;
    Procedure DoTreeTob( Range: String; T: Tob; nbNiveau: Integer );
    procedure DoClick( Sender: TObject );
    procedure DoSetNavigate( Sender: TObject );
    procedure PageChange( Sender: TObject );
    procedure ClickValBloque( Sender: TObject );
    procedure DlbClic( Sender: TObject );
    { EVI / Procedure pour corriger le redimensionnement des colonnes }
    procedure Resize;
    procedure TreeViewPopUp(Sender: TObject);
    procedure ImportDetCalend_OnClick(Sender: TObject);
    procedure CalendView_OnClick(Sender: TObject);
    procedure CalendDelete_OnClick(Sender: TObject);
    procedure CalendAllHisto_OnClick(Sender: TObject);
    procedure CalendAllType_OnClick(Sender: TObject);
    procedure CalendAllDelete_OnClick(Sender: TObject);
    procedure CalendHisto_OnClick(Sender: TObject);
    procedure CalendType_OnClick(Sender: TObject);
    procedure Refresh_OnClick(Sender: TObject);
    procedure Expand_OnClick(Sender: TObject);
    procedure ExpandAll_OnClick(Sender: TObject);
    procedure CollapseAll_OnClick(Sender: TObject);
    procedure CreateSousNiveauAuto_OnClick(Sender: TObject);
    procedure ModifSousNiveauAuto_OnClick(Sender: TObject);
    procedure SetImages (TV: TTreeView; TN: TTreeNode) ;
    procedure TV_OnExpanded(Sender: TObject; TN: TTreeNode);
    procedure TV_OnGetSelectedIndex(Sender: TObject; TN: TTreeNode);
    procedure AfterBParamClick(Sender: TObject);
    procedure ProcBtnEvolution(Sender: TObject);
    procedure ProcBSousNiveau(Sender: TObject);
    procedure ProcBDupliquer(Sender: TObject);
    procedure InsertNewSalary_OnClick(Sender: TObject);
    procedure ModifNodes_OnClick(Sender: TObject);
    procedure DoAfterLoadGrid( Sender: TObject );
    procedure DoOnAffecteCumul( Sender: TObject );

end ;

procedure GetTobNome( TNome: TOB; cProjet,codeSession: hString; iNivMax: Integer; cFils: hString );

Implementation

Uses Controls,Classes,graphics,sysutils,
     CstCommun,SaisieList,htb97,menus,ed_tools,HmsgBox,
     UDatamem,StrUtils,
     {$IFNDEF EAGLCLIENT} db,FE_main,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     {$else}MainEagl,{$ENDIF}
     BPCubeTMP,SynScriptBP,Uutil,BPBasic,QBPARBRE_TOM,
     BPFctArbre,BPFctSession,BPUtil;

procedure TOF_QUFSLARBRE.ClickValBloque(Sender: TObject);
begin
 OkModifValBloque:=true;
end;

procedure TOF_QUFSLARBRE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFSLARBRE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFSLARBRE.OnUpdate ;
begin
  Inherited ;

  //pour le rafraichissement d'un noeud après une duplication
  if (OkDuplicationRefresh) and (THTreeview(TF.LeTreeView).Selected<>nil) then
  begin
    TF.ReloadTreeNode(TF.LeTreeView.Selected);
    OkDuplicationRefresh:=false;
  end else TF.RefreshLignes;

 //si c'est une valeur bloquée
 //les valeurs qui permettent de modifier (montant et % d'évloution et valeur saisie)
 //sont écrites en rouges
 //sinon elles sont écrites en clWindowText
 if THCheckBox(GetControl('QBR_VALBLOQUE')).checked
  then
   begin
    //CA1
    THEdit(GetControl('QBR_EVOLPRCT1')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL1')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI1')).Font.Color:=clred;
    //CA2
    THEdit(GetControl('QBR_EVOLPRCT2')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL2')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI2')).Font.Color:=clred;
    //CA3
    THEdit(GetControl('QBR_EVOLPRCT3')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL3')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI3')).Font.Color:=clred;
    //CA4
    THEdit(GetControl('QBR_EVOLPRCT4')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL4')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI4')).Font.Color:=clred;
    //CA5
    THEdit(GetControl('QBR_EVOLPRCT5')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL5')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI5')).Font.Color:=clred;
    //CA6
    THEdit(GetControl('QBR_EVOLPRCT6')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL6')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI6')).Font.Color:=clred;
    //QTE
    THEdit(GetControl('QBR_EVOLQTEPRCT')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVALQTE')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISIQTE')).Font.Color:=clred;
    { EVI / Empêche la saisie si la case valeur bloquée est cochée }
    if ContextBP in [0,1,2] then THEdit(GetControl('PAGE')).Enabled:=False
    else
    begin
      THGroupBox(GetControl('BOX1')).Enabled:=False;
      THGroupBox(GetControl('BOX')).Enabled:=False;
      THGroupBox(GetControl('BOX3')).Enabled:=False;
      THGroupBox(GetControl('BOX4')).Enabled:=False;
      THGroupBox(GetControl('BOX5')).Enabled:=False;
      THGroupBox(GetControl('BOX6')).Enabled:=False;
      THGroupBox(GetControl('BOX7')).Enabled:=False;
    end;
   end
  else
   begin
    //CA1
    THEdit(GetControl('QBR_EVOLPRCT1')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVAL1')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISI1')).Font.Color:=clWindowText;
    //CA2
    THEdit(GetControl('QBR_EVOLPRCT2')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVAL2')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISI2')).Font.Color:=clWindowText;
    //CA3
    THEdit(GetControl('QBR_EVOLPRCT3')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVAL3')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISI3')).Font.Color:=clWindowText;
    //CA4
    THEdit(GetControl('QBR_EVOLPRCT4')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVAL4')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISI4')).Font.Color:=clWindowText;
    //CA5
    THEdit(GetControl('QBR_EVOLPRCT5')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVAL5')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISI5')).Font.Color:=clWindowText;
    //CA6
    THEdit(GetControl('QBR_EVOLPRCT6')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVAL6')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISI6')).Font.Color:=clWindowText;
    //QTE
    THEdit(GetControl('QBR_EVOLQTEPRCT')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVALQTE')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISIQTE')).Font.Color:=clWindowText;
    { EVI / Autorise la saisie si la case valeur bloquée est décochée }
    if ContextBP in [0,1,2] then THEdit(GetControl('PAGE')).Enabled:=true
    else
    begin
      THGroupBox(GetControl('BOX1')).Enabled:=true;
      THGroupBox(GetControl('BOX')).Enabled:=true;
      THGroupBox(GetControl('BOX3')).Enabled:=true;
      THGroupBox(GetControl('BOX4')).Enabled:=true;
      THGroupBox(GetControl('BOX5')).Enabled:=true;
      THGroupBox(GetControl('BOX6')).Enabled:=true;
      THGroupBox(GetControl('BOX7')).Enabled:=true;
    end;
   end;
end ;

procedure TOF_QUFSLARBRE.OnLoad ;
var i:integer;
begin
  Inherited ;
  if BPOkOrli
   then
    //-----------------> ORLI
    begin
      { EVI / ORLI : Onglets Grille de taille(9), CAPREV(10),QuantitéPrev(11) }
      // onglets visibles
      for i:=1 to 11 do
      THTabSheet(GetControl('TAB'+IntToStr(i))).tabVisible:=false;
      //si session de previson
      //on affiche que onglet qte
      if SessionInitCoeff(codeSession)
      then
      begin
        THCheckBox(GetControl('QBR_VALBLOQUE')).visible:=false;

        //si on se trouve au niveau du coeff
        //cad au dernier niveau de la session
        //on affiche l'onglet coeff et le bouton detail client
        if (niveauT=IntToStr(nivMax))
        then
        begin
          THTabSheet(GetControl('TAB8')).tabVisible:=true;
          TToolBarButton97(GetControl( 'BTNDETCLT' )).visible:=true;
        end
        else
        begin
          THTabSheet(GetControl('TAB8')).tabVisible:=false;
          TToolBarButton97(GetControl( 'BTNDETCLT' )).visible:=false;
        end;

        //suivant le parametre calcul le coeff en CA ou en qte
        //on affiche l'onglet correspondant

        if DonneParamS(ps_BPCoeffPerCAQte)='2'
        then
        begin
          THEdit(GetControl('QBR_HISTOCA')).visible:=false;
          THEdit(GetControl('QBR_REALISECA')).visible:=false;
          THEdit(GetControl('QBR_PREVUCA')).visible:=false;
          THEdit(GetControl('QBR_HISTO')).visible:=true;
          THEdit(GetControl('QBR_REALISE')).visible:=true;
          THEdit(GetControl('QBR_PREVU')).visible:=true;
        end
        else
        begin
          THEdit(GetControl('QBR_HISTOCA')).visible:=true;
          THEdit(GetControl('QBR_REALISECA')).visible:=true;
          THEdit(GetControl('QBR_PREVUCA')).visible:=true;
          THEdit(GetControl('QBR_HISTO')).visible:=false;
          THEdit(GetControl('QBR_REALISE')).visible:=false;
          THEdit(GetControl('QBR_PREVU')).visible:=false;
        end;

        if (OkSessionInitPrev)
        then
        begin
          THTabSheet(GetControl('TAB10')).tabVisible:=true;
          if (niveauT>IntToStr(nivMax))
          then THTabSheet(GetControl('TAB11')).tabVisible:=false
          else
          begin
            if DonneParamS(ps_BPCoeffPerCAQte)='2'
            then THTabSheet(GetControl('TAB11')).tabVisible:=false
            else THTabSheet(GetControl('TAB11')).tabVisible:=true;
          end;
        end
        else
        begin
          if DonneParamS(ps_BPCoeffPerCAQte)='2'
          then
          begin
            THTabSheet(GetControl('TAB11')).tabVisible:=false;
            THTabSheet(GetControl('TAB10')).tabVisible:=true;
          end
          else
          begin
            THTabSheet(GetControl('TAB11')).tabVisible:=true;
            THTabSheet(GetControl('TAB10')).tabVisible:=false;
          end;
        end;

      end
      else
      begin
        THCheckBox(GetControl('QBR_VALBLOQUE')).visible:=true;
        for i:=0 to 10 do
        if TabValAff[i]<>''
        then
        begin
          THTabSheet(GetControl(TabValAff[i])).tabVisible:=true;
          THTabSheet(GetControl(TabValAff[i])).caption:=TabLibValAff[i];
        end;
      end;
      PageChange(self);
    end
    //ORLI <-----------------
    else
    begin
      TToolBarButton97(GetControl( 'BTNDETCLT' )).visible:=false;
      Case ContextBP of
        0,1,2 : begin
                case ValeurAffiche of
                  1 : THTabSheet(GetControl('TAB1')).tabVisible:=true;
                  2 : THTabSheet(GetControl('TAB2')).tabVisible:=true;
                  3 : THTabSheet(GetControl('TAB3')).tabVisible:=true;
                  4 : THTabSheet(GetControl('TAB4')).tabVisible:=true;
                  5 : THTabSheet(GetControl('TAB5')).tabVisible:=true;
                  6 : THTabSheet(GetControl('TAB6')).tabVisible:=true;
                  7 : THTabSheet(GetControl('TAB7')).tabVisible:=true;
                end
              end;
        3 : begin  //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
              //Renomme les champs Quantité et cache le champ "Retenue"
              SetControlText('TQBR_QTEREF',TraduireMemoire('Montant historique'));
              SetControlText('TQBI_CORRECTIFVAL1',TraduireMemoire('Montant d''évolution'));
              SetControlVisible('TQBR_QTERETENUE',False);
              SetControlVisible('QBR_QTERETENUE',False);

              for i:= 1 to NbValAff do
              begin
                if NbValAff=1 then
                begin
                   THTabSheet(GetControl('TAB2')).tabVisible:=true ;
                   THTabSheet(GetControl('TAB1')).tabVisible:=false ;
                end
                else
                begin
                  if i = 1 then THTabSheet(GetControl('TAB2')).tabVisible:=true
                  else if i = 2 then THTabSheet(GetControl('TAB1')).tabVisible:=true
                  else THTabSheet(GetControl('TAB'+IntToStr(i))).tabVisible:=true;
                end;
              end;

              if NbValAff > 1 then
              begin
                TPageControl(getControl( 'PAGE')).Pages[1].PageIndex := 0;
                TPageControl(getControl( 'PAGE')).ActivePage := THTabSheet(GetControl('TAB2'));
              end;
            end;
    end;
    PageChange(self);
  end;
  OkDuplicationRefresh:=false;
  if THCheckBox(GetControl('QBR_VALBLOQUE')).checked
  then
  begin
    //CA1
    THEdit(GetControl('QBR_EVOLPRCT1')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL1')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI1')).Font.Color:=clred;
    //CA2
    THEdit(GetControl('QBR_EVOLPRCT2')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL2')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI2')).Font.Color:=clred;
    //CA3
    THEdit(GetControl('QBR_EVOLPRCT3')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL3')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI3')).Font.Color:=clred;
    //CA4
    THEdit(GetControl('QBR_EVOLPRCT4')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL4')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI4')).Font.Color:=clred;
    //CA5
    THEdit(GetControl('QBR_EVOLPRCT5')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL5')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI5')).Font.Color:=clred;
    //CA6
    THEdit(GetControl('QBR_EVOLPRCT6')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL6')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI6')).Font.Color:=clred;
    //QTE
    THEdit(GetControl('QBR_EVOLQTEPRCT')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVALQTE')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISIQTE')).Font.Color:=clred;
    { EVI / Empêche la saisie si la case valeur bloquée est cochée }
    if ContextBP in [0,1,2] then THEdit(GetControl('PAGE')).Enabled:=False
    else
    begin
      THGroupBox(GetControl('BOX1')).Enabled:=False;
      THGroupBox(GetControl('BOX')).Enabled:=False;
      THGroupBox(GetControl('BOX3')).Enabled:=False;
      THGroupBox(GetControl('BOX4')).Enabled:=False;
      THGroupBox(GetControl('BOX5')).Enabled:=False;
      THGroupBox(GetControl('BOX6')).Enabled:=False;
      THGroupBox(GetControl('BOX7')).Enabled:=False;
    end;
  end
  else
  begin
    //CA1
    THEdit(GetControl('QBR_EVOLPRCT1')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVAL1')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISI1')).Font.Color:=clWindowText;
    //CA2
    THEdit(GetControl('QBR_EVOLPRCT2')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVAL2')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISI2')).Font.Color:=clWindowText;
    //CA3
    THEdit(GetControl('QBR_EVOLPRCT3')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVAL3')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISI3')).Font.Color:=clWindowText;
    //CA4
    THEdit(GetControl('QBR_EVOLPRCT4')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVAL4')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISI4')).Font.Color:=clWindowText;
    //CA5
    THEdit(GetControl('QBR_EVOLPRCT5')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVAL5')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISI5')).Font.Color:=clWindowText;
    //CA6
    THEdit(GetControl('QBR_EVOLPRCT6')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVAL6')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISI6')).Font.Color:=clWindowText;
    //QTE
    THEdit(GetControl('QBR_EVOLQTEPRCT')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTEVOLVALQTE')).Font.Color:=clWindowText;
    THEdit(GetControl('EDTSAISIQTE')).Font.Color:=clWindowText;
    { EVI / Autorise la saisie si la case valeur bloquée est décochée }
    if ContextBP in [0,1,2] then THEdit(GetControl('PAGE')).Enabled:=true
    else
    begin
      THGroupBox(GetControl('BOX1')).Enabled:=true;
      THGroupBox(GetControl('BOX')).Enabled:=true;
      THGroupBox(GetControl('BOX3')).Enabled:=true;
      THGroupBox(GetControl('BOX4')).Enabled:=true;
      THGroupBox(GetControl('BOX5')).Enabled:=true;
      THGroupBox(GetControl('BOX6')).Enabled:=true;
      THGroupBox(GetControl('BOX7')).Enabled:=true;
    end;
  end;

  Img0 := TImage( GetControl( 'TREEICON1' ) );
  Img1 := TImage( GetControl( 'TREEICON2' ) );
  Img2 := TImage( GetControl( 'TREEICON3' ) );
  Img3 := TImage( GetControl( 'TREEICON4' ) );
  Img4 := TImage( GetControl( 'TREEICON5' ) );
  ImgSel := TCustomImageList.Create( TV );
  ImgSel.Add( Img0.Picture.Bitmap, nil );
  ImgSel.Add( Img1.Picture.Bitmap, nil );
  ImgSel.Add( Img2.Picture.Bitmap, nil );
  ImgSel.Add( Img3.Picture.Bitmap, nil );
  ImgSel.Add( Img4.Picture.Bitmap, nil );
  TV.Images := ImgSel;

end ;

procedure TOF_QUFSLARBRE.OnArgument (S : String ) ;
var TabLibelle : array[0..7] of hString;
Q:Tquery;
k,i:integer;
DateDeb,DateFin:TDateTime;
begin
  Inherited ;
  OldCodeAxe:='0';
  NodeDelete:=0;
  TF := TFSaisieList(Ecran).LeFiltre;
  TF.OnAfterLoadGrid := DoAfterLoadGrid;
  TF.OnTreeTob := DoTreeTob;

  TV := THTreeView( GetControl( 'TreeEntete' ) );
  TV.OnExpanded := TV_OnExpanded;
  TV.OnGetSelectedIndex := TV_OnGetSelectedIndex;
  TF.OnSetNavigate := DoSetNavigate;

  TPopUpMenu(GetControl('POPTREEVIEWMENU')).OnPopUp := TreeViewPopUp;
  TToolBarButton97(GetControl( 'BSOUSNIVEAU' )).onClick := DoClick;
  TToolBarButton97(GetControl( 'BDUPLIQUER' )).onClick := DoClick;
  //Accès interdit pour la duplication avec Historique A REVOIR
  //TToolBarButton97(GetControl( 'BDUPLIQUERREF' )).onClick := DoClick;
  TToolBarButton97(GetControl( 'BDUPLIQUERREF' )).Visible := False;
  TToolBarButton97(GetControl( 'BDELETE' )).onClick := DoClick;
  TToolBarButton97(GetControl( 'BINSERT' )).onClick := DoClick;
  TToolBarButton97(GetControl( 'BTNEVOLUTION' )).onClick := DoClick;
  TToolBarButton97(GetControl( 'BTNDETCLT' )).onClick := DoClick;
  TToolBarButton97(GetControl( 'BTNCALREP' )).onClick := DoClick;
  TToolBarButton97(GetControl( 'FakeBParamListe')).onClick := AfterBParamClick;
  TToolBarButton97(GetControl( 'BTNCSUBNIVAUTO' )).onClick := CreateSousNiveauAuto_OnClick;
  TToolBarButton97(GetControl( 'BTNMSUBNIVAUTO' )).onClick := ModifSousNiveauAuto_OnClick;
  TToolBarButton97(GetControl( 'BTNNEWSALPAIE' )).onClick := InsertNewSalary_OnClick;
  TToolBarButton97(GetControl( 'BTNMODIFNODES' )).onClick := ModifNodes_OnClick;
  if ContextBP=1 then THGrid(GetControl( 'FLISTE' )).OnDblClick :=DlbClic;
  THPageControl2(GetControl('PAGE' )).OnChange:=PageChange;
  THCheckBox(GetControl('QBR_VALBLOQUE')).OnClick:=ClickValBloque;

  consult:=TrouveArgument(S,'CONSULT','');

  TFsaisieList( ecran ).SavePosition := true;
  codeSession:=TrouveArgument(S,'SESSION','');
  nivMax:=ChercheNivMax(codeSession);

  ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);
  sousNivOk:=false;

  OkSessionObj:=OkSessionObjectif(codeSession);
  OkSessionInitCoeff:=SessionInitCoeff(codeSession);
  OkSessionInitPrev:=SessionInitPrev(codeSession);
  ValeurAffiche:=DonneValeurAffiche(codesession);
  BPInitialise:=SessionBPInitialise(codeSession);
  if BPInitialise = '' then BPInitialise := '0';
  OkMaillageJour:=SessionMaillageJour(codeSession);

  if not BPOkOrli then
  begin

    Case ContextBP of
      2 : begin //Compta
            THTabSheet(GetControl('TAB1')).Caption := 'Débit - Crédit';
            THTabSheet(GetControl('TAB3')).Caption := 'Crédit - Débit';
          end;
      3 : begin //Paie
            LibValAff(codeSession,TabLibelle);
            NbValAff := StrToInt(TabLibelle[0]);
            for i:=1 to NbValAff do
            begin               //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
              if i = 1 then THTabSheet(GetControl('TAB2')).Caption := TabLibelle[i]
              else if i = 2 then THTabSheet(GetControl('TAB1')).Caption := TabLibelle[i]
              else THTabSheet(GetControl('TAB'+IntToStr(i))).Caption := TabLibelle[i]
            end;
            Q:=OpenSQL('SELECT QBS_VALAFFH1,QBS_VALAFFH2,QBS_VALAFFH3,QBS_VALAFFH4,QBS_VALAFFH5,QBS_VALAFFH6,QBS_VALAFFH7'+
                      ' FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+CodeSession+'"',true);
            if not Q.Eof then
            begin
              for i:=1 to NbValAff do
              begin               //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
                if i = 1 then TabValues[i+1] := Q.Fields[i-1].AsString
                else if i = 2 then TabValues[i-1] := Q.Fields[i-1].AsString
                else TabValues[i] := Q.Fields[i-1].AsString
              end;
            end;
            Ferme(Q);

            if TabValues[1]='ETP' then
            begin
              TabRefSQLValues[1]:='QBR_REF1';
              TabRefSQLValues[8]:='QBR_OP1';
            end;
            if TabValues[2]='ETP' then
            begin
              TabRefSQLValues[3]:='QBR_QTEREF';
              TabRefSQLValues[9]:='QBR_QTEC';
            end;
            for k:=3 to 7 do
            begin
              if TabValues[k]='ETP' then
              begin
                TabRefSQLValues[k]:='QBR_REF'+IntToStr(k-1);
                TabRefSQLValues[k+7]:='QBR_OP'+IntToStr(k-1);
              end;
            end;
          end;
    end;

    Case ContextBP of
      0,1,2 : begin //Mode-GC-Compta
              case ValeurAffiche of
                1 : begin
                      TF.LaGridListe:='QULBPARBRECA1';
                      TabValAff[0]:='PTC';
                      THTabSheet(GetControl('TAB1')).tabVisible:=true;
                    end;
                2 : begin
                      TF.LaGridListe:='QULBPARBREQTE';
                      TabValAff[0]:='QTE';
                      THTabSheet(GetControl('TAB2')).tabVisible:=true;
                    end;
                3 : begin
                      TF.LaGridListe:='QULBPARBRECA2';
                      TabValAff[0]:='PHT';
                      THTabSheet(GetControl('TAB3')).tabVisible:=true;
                    end;
                4 : begin
                      TF.LaGridListe:='QULBPARBRECA3';
                      TabValAff[0]:='UTC';
                      THTabSheet(GetControl('TAB4')).tabVisible:=true;
                    end;
                5 : begin
                      TF.LaGridListe:='QULBPARBRECA4';
                      TabValAff[0]:='UHT';
                      THTabSheet(GetControl('TAB5')).tabVisible:=true;
                    end;
                6 : begin
                      TF.LaGridListe:='QULBPARBRECA5';
                      TabValAff[0]:='PAH';
                      THTabSheet(GetControl('TAB6')).tabVisible:=true;
                    end;
                7 : begin
                      TF.LaGridListe:='QULBPARBRECA6';
                      TabValAff[0]:='MAR';
                      THTabSheet(GetControl('TAB7')).tabVisible:=true;
                    end;
                 end;
            end ;
      3 :   TF.LaGridListe:='QULBPARBREPAIE'+IntToStr(NbValAff);

    end
  end
  else //-----------------> ORLI
  RechercheTabAff(TabValAff,TabLibValAff);
      //ORLI <-----------------

  ChangeListeLibelleColonnes(OkSessionObj,codeSession,TabValAff,TabLibValAff);

  if ContextBP=3 then
  begin
    dateDebFinSesssion(codeSession,DateDeb,DateFin);
    NbIntervalle:=DonneNbIntervalleMailleDateDebDateFin('4',DateDeb,DateFin);
  end;
end ;

procedure TOF_QUFSLARBRE.OnClose ;
begin
  Inherited ;
  ControleNoeudANoeudPere(codeSession);
end ;

procedure GetTobNome( TNome: TOB; cProjet,codeSession: hString; iNivMax: Integer; cFils: hString );
var
  T: Tob;

  function LoadTobFromSql(Sql: hString; t:Tob; lAppend: boolean = false): boolean;
  var
    Q: tQuery;
  begin
    Result := false;
    Q := OpenSql(Sql, True);
     try
       if not Q.Eof then
        begin
          t.loadDetailDB('', '', '', Q, lAppend);
           Result := true;
        end;
     finally
      Ferme(Q);
     end;
  end;

  function FindLeNiveau( cQuoi: hString; iNiv: integer; TPere: TOB; cFils: hString= ''): Boolean;
  var TFils, Tnew: tob;
      i,j,k: integer;
      AxeSQL,libValAxe:hString;
  begin
    inc(iNiv);
    if iNiv>iNivMax then
      begin
      Result := ExisteSQL('SELECT (1) FROM QBPARBRE WHERE  '+
                          'QBR_CODESESSION="'+codeSession+
                          '" AND QBR_NUMNOEUDPERE="'+cquoi+'"' );
      exit;
      end;
    TFils := Tob.Create( 'Virtuelle ',nil,-1 );
    if cFils<>'' then //PMJ
      { EVI / Ajout de QBR_VALBLOQUE,QBR_NIVEAU dans la TOB }
      LoadTobFromSql('SELECT QBR_NUMNOEUDPERE,QBR_CODESESSION,QBR_NUMNOEUD,'+
                     'QBR_VALEURAXE,QBR_CODEAXE,QBR_DATEDELAI,QBR_DEVISE,QBR_VALBLOQUE,QBR_NIVEAU,'+
                     'QBR_VALAXENIV1,QBR_VALAXENIV2,QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,'+
                     'QBR_VALAXENIV6,QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,QBR_VALAXENIV10'+
                     ' FROM QBPARBRE '+
                     'WHERE QBR_CODESESSION="'+codeSession+
                     '" AND QBR_NUMNOEUD="'+cFils+'" ORDER BY QBR_DATEDELAI,QBR_VALEURAXE',TFils )
    else
      LoadTobFromSql('SELECT QBR_NUMNOEUDPERE,QBR_CODESESSION,QBR_NUMNOEUD,'+
                     'QBR_VALEURAXE,QBR_CODEAXE,QBR_DATEDELAI,QBR_DEVISE,QBR_VALBLOQUE,QBR_NIVEAU,'+
                     'QBR_VALAXENIV1,QBR_VALAXENIV2,QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,'+
                     'QBR_VALAXENIV6,QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,QBR_VALAXENIV10'+
                     ' FROM QBPARBRE '+
                     'WHERE QBR_CODESESSION="'+codeSession+
                     '" AND QBR_NUMNOEUDPERE="'+cprojet+'" ORDER BY QBR_DATEDELAI,QBR_VALEURAXE',TFils );
    for i := 0 to TFils.Detail.Count - 1 do
      begin
      TNew := Tob.Create ( 'QBPARBRED', TPere, -1 );
      TNew.PutValue('QBG_CODESESSION',TFils.Detail[i].GetValue('QBR_CODESESSION'));
      TNew.PutValue('QBG_NUMNOEUD',TFils.Detail[i].GetInteger('QBR_NUMNOEUD'));
      TNew.PutValue('QBG_NUMNOEUDPERE',TFils.Detail[i].GetInteger('QBR_NUMNOEUDPERE'));
      { EVI / Valeur par défaut pour axe sans libellé }
      if TFils.Detail[i].GetString('QBR_VALEURAXE')='' then  TNew.PutValue('QBG_VALEURAXE','[---]')
      else TNew.PutValue('QBG_VALEURAXE',TFils.Detail[i].GetString('QBR_VALEURAXE'));
      TNew.PutValue('QBG_CODEAXE',TFils.Detail[i].GetString('QBR_CODEAXE'));
      TNew.PutValue('QBG_DEVISE',TFils.Detail[i].GetString('QBR_DEVISE'));
      libValAxe:=DonneLibelleValeurAxe(TFils.Detail[i].GetString('QBR_CODEAXE'),
                                       TFils.Detail[i].GetString('QBR_VALEURAXE'));
      TNew.PutValue('QBG_LIBVALAXE',libValAxe);
      Tnew.AddChampSupValeur('CLELIEN',CalcCleLien( 'QBPARBRE', TFils.Detail[i] ) );
      if FindLeNiveau( TFils.Detail[i].GetValue('QBR_NUMNOEUD'),iNiv,TNome ) then
      TNew.AddChampSupValeur('SOUSNIVEAU',True) else TNew.AddChampSupValeur('SOUSNIVEAU',False);
      TNew.PutValue('QBG_VALBLOQUE',TFils.Detail[i].GetString('QBR_VALBLOQUE'));
      TNew.PutValue('QBG_NIVEAU',TFils.Detail[i].GetString('QBR_NIVEAU'));

      j:=1; AxeSQL := '';

      While j < TFils.Detail[i].GetInteger('QBR_NIVEAU') do
      begin
        if TFils.Detail[i].GetString('QBR_VALAXENIV'+IntToStr(j)) = '' then
        begin
          TNew.PutValue('QBG_VALAXENIV'+IntToStr(j),'[---]');
          AxeSQL := AxeSQL + ' AND QBE_VALAXENIV'+IntToStr(j)+'="[---]"';
        end
        else
        begin
          TNew.PutValue('QBG_VALAXENIV'+IntToStr(j),TFils.Detail[i].GetString('QBR_VALAXENIV'+IntToStr(j)));
          AxeSQL := AxeSQL + ' AND QBE_VALAXENIV'+IntToStr(j)+'="'+TFils.Detail[i].GetString('QBR_VALAXENIV'+IntToStr(j))+'"';
        end;
        j:=j+1;
      end;

      if j<>11 then
      begin
        if TFils.Detail[i].GetString('QBR_VALEURAXE') = ''
        then AxeSQL := AxeSQL + ' AND QBE_VALAXENIV'+IntToStr(j)+'="[---]"'
        else AxeSQL := AxeSQL + ' AND QBE_VALAXENIV'+IntToStr(j)+'="'+TFils.Detail[i].GetString('QBR_VALEURAXE')+'"';
      end;

      //Complète la requête avec les axes vides
      for k:= (j+1) to 10 do AxeSQL :=  AxeSQL + ' AND QBE_VALAXENIV'+IntToStr(k)+'=""';

      if ExisteSQL('SELECT QBE_CALENDREP, QBE_VALAXENIV1, QBE_VALAXENIV2, QBE_VALAXENIV3 , QBE_VALAXENIV4,'+
      ' QBE_VALAXENIV5, QBE_VALAXENIV6, QBE_VALAXENIV7, QBE_VALAXENIV8, QBE_VALAXENIV9, QBE_VALAXENIV10'+
      ' FROM QBPDETCALENDREP WHERE QBE_CALENDREP="'+CodeSession+'"'+AxeSQL+
      ' GROUP BY QBE_CALENDREP, QBE_VALAXENIV1, QBE_VALAXENIV2, QBE_VALAXENIV3, QBE_VALAXENIV4, QBE_VALAXENIV5, '+
      ' QBE_VALAXENIV6, QBE_VALAXENIV7, QBE_VALAXENIV8, QBE_VALAXENIV9, QBE_VALAXENIV10')
      then Tnew.AddChampSupValeur('CALEND','X' )
      else Tnew.AddChampSupValeur('CALEND','-' );
    end;
    Result := TFils.Detail.Count>0;
    TFils.Free;
  end;

begin
  if TNome.NomTable = 'QBPARBRED' then
    begin
     TNome.PutValue('QBG_CODESESSION',codeSession);
     TNome.PutValue('QBG_NUMNOEUDPERE',cProjet);
    end
   else
    begin
     TNome.PutValue('QBR_CODESESSION',codeSession);
     TNome.PutValue('QBR_NUMNOEUDPERE',cProjet);
     TNome.PutValue('QBR_NUMNOEUD',cfils);
    end;
  T := Tob.Create( 'Virtuelle ',nil,-1 );

  LoadTobFromSql('SELECT QBR_VALEURAXE FROM QBPARBRE '+
                 'WHERE QBR_CODESESSION="'+codeSession+
                 '" AND QBR_NUMNOEUD="'+cfils+'"',T );
  if T.Detail.count>0 then
  begin
    if TNome.NomTable = 'QBPARBRED'
    then TNome.PutValue('QBG_VALEURAXE',T.Detail[0].GetValue('QBR_VALEURAXE'))
    else TNome.PutValue('QBR_VALEURAXE',T.Detail[0].GetValue('QBR_VALEURAXE'));
  end
  { EVI / Titre dans le premier noeud de l'arbre }
  else TNome.PutValue('QBG_VALEURAXE',codeSession);

  T.Free;

  if FindLeNiveau( cProjet, 0, TNome, cFils ) then
    TNome.AddChampSupValeur('SOUSNIVEAU',True)
  else
    TNome.AddChampSupValeur('SOUSNIVEAU',False);

end;

procedure TOF_QUFSLARBRE.DoTreeTob(Range: String; T: Tob;
  nbNiveau: Integer);
var cProjet,cFils,codeAxe: hString;
begin
 cProjet := '';
  cFils := '';
  if range<>''
   then
    begin
     codeSession:=TrouveArgument(range,'SESSION','');
     cProjet:=TrouveArgument(range,'NUMNOEUDPERE','0');
    end
  else if T.NomTable = 'QBPARBRED' then
    begin
    cProjet := T.GetValue('QBG_NUMNOEUD');
    end
  else if T.NomTable = 'QBPARBRE' then
    begin
    cProjet := T.GetValue('QBR_NUMNOEUDPERE');
    cFils := T.GetValue('QBR_NUMNOEUD');
    codeAxe := T.GetValue('QBR_CODEAXE');
    end;

  GetTobNome( T, cProjet,codeSession, nbNiveau, cFils);
end;

procedure TOF_QUFSLARBRE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFSLARBRE.OnCancel () ;
begin
  Inherited ;
  sousNivOk:=false;
end ;

procedure SuppressionNiveauSuivant(numnoeud,codesession:hString);
var Q:TQuery;
    numnoeudFils:hString;
begin
 Q:=MOpenSql('SELECT QBR_NUMNOEUD FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUDPERE="'+numnoeud+'"','QUFSLARBRE_TOF (SuppressionNiveauSuivant).',true);
 while not Q.eof do
 begin
   numnoeudFils:=Q.fields[0].asString;
   //suppression du niveau courant "niveau"
   MExecuteSql('DELETE FROM QBPARBRE '+
            ' WHERE QBR_CODESESSION="'+codeSession+
            '" AND QBR_NUMNOEUD="'+numnoeudFils+'"',
            'QUFSLARBRE_TOF (SuppressionNiveauSuivant).');
   SuppressionNiveauSuivant(numnoeudFils,codeSession);
   Q.next;
 end;
 ferme(Q);
end;

procedure TOF_QUFSLARBRE.DlbClic(Sender: TObject);
begin
 if TOB(TV.Selected.Data).FieldExists('CALEND') then
 begin
   if TOB(TV.Selected.Data).GetValue('CALEND') = 'X' then
   DoClick(  TToolBarButton97(GetControl( 'BTNCALREP' )));
 end;
end;

procedure TOF_QUFSLARBRE.ProcBtnEvolution(Sender: TObject);
begin
  DoClick(  TToolBarButton97(GetControl( 'BTNEVOLUTION' )));
end;

procedure TOF_QUFSLARBRE.ProcBSousNiveau(Sender: TObject);
begin
  DoClick(  TToolBarButton97(GetControl( 'BSOUSNIVEAU' )));
end;

procedure TOF_QUFSLARBRE.ProcBDupliquer(Sender: TObject);
begin
  DoClick(  TToolBarButton97(GetControl( 'bDupliquer' )));
end;

procedure TOF_QUFSLARBRE.DoClick(Sender: TObject);
var nivI,i,OngletActif:integer;
    noeud,numnoeudpere,retour:hString;
    AncQte,ancCA1,ancCA2,ancCA3,ancCA4,ancCA5,ancCA6,ancPrevu,ancRet,AncCaRetenu:double;
    Prevu,Histo:double;
    codeCol,codeTitre,CoeffRetenu,Axes:hString;
    DateDebC,DateFinC:TDateTime;
begin
  OkRepriseHisto:=true;
  sousNivOk:=false;
  Prevu:=0;
  Histo:=0;

  if TControl( Sender ).name = 'BSOUSNIVEAU'
  then
  begin
    numnoeudT:=IntToStr(BPIncrementenumNoeud(codeSession));
    numnoeudpereT:=TF.GetValue('QBR_NUMNOEUD');
    { EVI / Correction violation d'accès si BSOUSNIVEAU sur premier noeud de l'arbre }
    if Tob(TV.Selected.data).GetValue('QBG_NUMNOEUD') = 0 then nivI := 0
    else nivI:=VALEURI(TF.GetValue('QBR_NIVEAU'));
    niveauT:=IntToStr(nivI+1);
    codeaxeT:=DonneCodeAxeNiv(codeSession,niveauT);
    if SessionDelai(codeSession) and (valeuri(niveauT)=nivMax+1) then codeaxeT:='DELAI';

    THLabel(GetControl( 'LBLCODEAXE' )).Caption:=codeaxeT;
    THLabel(GetControl( 'TQBR_VALEURAXE' )).Caption:=codeaxeT;
    if not BPOkOrli
     then
      begin
       THLabel(GetControl( 'LBLCODEAXE' )).Caption:=DonneLibelleCAxe('','',codeaxeT);
       THLabel(GetControl( 'TQBR_VALEURAXE' )).Caption:=DonneLibelleCAxe('','',codeaxeT);
      end;

    ValAxeNiv1T:=TF.GetValue('QBR_VALAXENIV1');
    ValAxeNiv2T:=TF.GetValue('QBR_VALAXENIV2');
    ValAxeNiv3T:=TF.GetValue('QBR_VALAXENIV3');
    ValAxeNiv4T:=TF.GetValue('QBR_VALAXENIV4');
    ValAxeNiv5T:=TF.GetValue('QBR_VALAXENIV5');
    ValAxeNiv6T:=TF.GetValue('QBR_VALAXENIV6');
    ValAxeNiv7T:=TF.GetValue('QBR_VALAXENIV7');
    ValAxeNiv8T:=TF.GetValue('QBR_VALAXENIV8');
    ValAxeNiv9T:=TF.GetValue('QBR_VALAXENIV9');
    ValAxeNiv10T:=TF.GetValue('QBR_VALAXENIV10');
    case nivI of
     1 : ValAxeNiv1T:=TF.GetValue('QBR_VALEURAXE');
     2 : ValAxeNiv2T:=TF.GetValue('QBR_VALEURAXE');
     3 : ValAxeNiv3T:=TF.GetValue('QBR_VALEURAXE');
     4 : ValAxeNiv4T:=TF.GetValue('QBR_VALEURAXE');
     5 : ValAxeNiv5T:=TF.GetValue('QBR_VALEURAXE');
     6 : ValAxeNiv6T:=TF.GetValue('QBR_VALEURAXE');
     7 : ValAxeNiv7T:=TF.GetValue('QBR_VALEURAXE');
     8 : ValAxeNiv8T:=TF.GetValue('QBR_VALEURAXE');
     9 : ValAxeNiv9T:=TF.GetValue('QBR_VALEURAXE');
     10 : ValAxeNiv10T:=TF.GetValue('QBR_VALEURAXE');
    end;

    sousNivOk:=true;
    if niveauT<>'1'
    then
    begin
      TF.CreateSousNiveau('QBG_CODESESSION='+TF.GetValue('QBR_CODESESSION')+';QBG_NUMNOEUDPERE='+hString(TF.GetValue('QBR_NUMNOEUD')));
      TToolBarButton97(GetControl('BPost')).Enabled:=false;
    end
    else
    begin
      numnoeudpereT:='0';
      TF.CreateSousNiveau('QBG_CODESESSION='+codeSession+';QBG_NUMNOEUDPERE='+'0');
      TToolBarButton97(GetControl('BPost')).Enabled:=false;
    end;
    OkDuplicationRefresh:=true;
  end
  else
  if TControl( Sender ).name = 'bDupliquer'
  then
  begin
    //Protection si tentative de duplication de l'entête de l'arbre
    if Tob(TV.Selected.data).GetValue('QBG_NIVEAU')=0 then exit;
    OkRepriseHisto:=false;
    numNoeudDup:=TF.GetValue('QBR_NUMNOEUD');
    numnoeudT:=IntToStr(BPIncrementenumNoeud(codeSession));
    noeud:=numnoeudT;
    TF.Duplicate;
    OkDuplication:=true;
    OkDuplicationRefresh:=true;
  end
  else
  if TControl( Sender ).name = 'BDUPLIQUERREF'
  then
  begin
    OkRepriseHisto:=true;
    numNoeudDup:=TF.GetValue('QBR_NUMNOEUD');
    numnoeudT:=IntToStr(BPIncrementenumNoeud(codeSession));
    noeud:=numnoeudT;
    TF.Duplicate;
    OkDuplication:=true;
    OkDuplicationRefresh:=true;
  end
  else
  if TControl( Sender ).name = 'BDelete'
  then
  begin
    //Empêche la suppression d'un noued avec historique
    case ContextBP of
     0,1,2 : begin
             case ValeurAffiche of
              1 : Histo := TF.GetValue('QBR_REF1');
              2 : Histo := TF.GetValue('QBR_QTEREF');
              3 : Histo := TF.GetValue('QBR_REF2');
              4 : Histo := TF.GetValue('QBR_REF3');
              5 : Histo := TF.GetValue('QBR_REF4');
              6 : Histo := TF.GetValue('QBR_REF5');
              7 : Histo := TF.GetValue('QBR_REF6');
             end;
           end;
     3 : begin
           if StrToInt(AnsiRightStr(TabActive,1)) = 1 then Histo := TF.GetValue('QBR_REF1')
           else if StrToInt(AnsiRightStr(TabActive,1)) = 2 then Histo := TF.GetValue('QBR_QTEREF')
           else Histo := TF.GetValue('QBR_REF'+IntToStr(StrToInt(AnsiRightStr(TabActive,1))-1));
         end;
    end;

    if Histo = 0 then
    begin
      NodeDelete:=2;
      noeud:=TF.GetValue('QBR_NUMNOEUD');
      numnoeudpere:=TF.GetValue('QBR_NUMNOEUDPERE');
      ArbreTotalNiv(codeSession,numNoeudPere,AncQte,ancCA1,ancCA2,
                    ancCA3,ancCA4,ancCA5,ancCA6,ancPrevu,ancRet,AncCaRetenu);
      OkDelete:=false;
      if OkModifNoeudArbre(codeSession,noeud,numnoeudpere,'0','1',nivMax,OkModifValBloque)
      then TF.Delete
      else PGIINFO('Valeur non modifiable (blocage des valeurs des autres niveaux).');
      if OkDelete
      then
      begin
        beginTrans;
        SuppressionNiveauSuivant(noeud,codeSession);
        //mise à jour du niveau courant et des niveaux précedants
        {ArbreTotalNiv(codeSession,numNoeudPere,totQteVal,totVal,TotCAVal2,
                      TotCAVal3,TotCAVal4,TotCAVal5,TotCAVal6,TotalPrevu,TotalRet);
        nivI:=VALEURI(TF.GetValue('QBR_NIVEAU'));
        MiseAjourNiveauPrecedent(true,codeSession,noeud,numnoeudpere,nivI,
                                 totVal,TotQteVal,
                                 TotCAVal2,TotCAVal3,TotCAVal4,TotCAVal5,TotCAVal6,
                                 totVal-ancCA1,totQteVal-ancQte,
                                 TotCAVal2-ancCA2,TotCAVal3-ancCA3,
                                 TotCAVal4-ancCA4,TotCAVal5-ancCA5,TotCAVal6-ancCA6,
                                 TotalPrevu,TotalRet,'',false);
        }

        MiseAJourPrctArbre(codeSession,true,true);
        commitTrans;
        TF.RefreshLignes;
      end;
    end else PGIINFO('Seules les valeurs sans données historique peuvent être supprimées.');
  end
  else
  if TControl( Sender ).name = 'Binsert' then
  begin
    numnoeudT:=IntToStr(BPIncrementenumNoeud(codeSession));
    numnoeudpereT:=TF.GetValue('QBR_NUMNOEUDPERE');
    niveauT:=TF.GetValue('QBR_NIVEAU');
    codeaxeT:=TF.GetValue('QBR_CODEAXE');
    TF.Insert;
    OkDuplicationRefresh:=true;
  end
  else
  if TControl( Sender ).name = 'BTNEVOLUTION'
  then
  begin
    numnoeudT:=TF.GetValue('QBR_NUMNOEUD');
    numnoeudpereT:=TF.GetValue('QBR_NUMNOEUDPERE');
    niveauT:=TF.GetValue('QBR_NIVEAU');
    codeaxeT:=TF.GetValue('QBR_CODEAXE');
    OngletActif := StrToInt(RightStr(TPageControl(getControl( 'PAGE')).ActivePage.Name,1));
    if not BPOkOrli then
    retour:=aglLanceFiche('Q', 'QUFVBPEVOLUTION', '', '',
                          'CODEAXE='+DonneLibelleAxe(codeaxeT)+';VALEURAXE='+
                          TF.GetValue('QBR_VALEURAXE')+
                          ';CODEAXEPREC='+DonneLibelleAxe(TabCodeAxe[VALEURI(niveauT)]))
    else //-----------------> ORLI
    retour:=aglLanceFiche('Q', 'QUFVBPEVOLUTION', '', '',
                          'CODEAXE='+codeaxeT+';VALEURAXE='+
                          TF.GetValue('QBR_VALEURAXE')+
                          ';CODEAXEPREC='+TabCodeAxe[VALEURI(niveauT)]);
         //ORLI <-----------------
    if retour<>'' then EvolutionSurUneValeurAxe(retour,codeSession,codeaxeT,TF.GetValue('QBR_VALEURAXE'),OngletActif);
    TF.RefreshLignes;
  end
  else
  if TControl( Sender ).name = 'BTNDETCLT'
  then
  begin
    codeCol:=' QBU_CODESESSION="'+codeSession+'" ';
    CoeffRetenu:=TF.GetValue('QBR_COEFFRETENU');
    for i:=1 to NivMax-1 do
    begin
      CodeCol:=codeCol+' AND QBU_VALAXECU'+IntToStr(i)+'="'+TF.GetValue('QBR_VALAXENIV'+IntToStr(i))+'" ';
      codeTitre:=codeTitre+TabCodeAxe[i+1]+' :'+TF.GetValue('QBR_VALAXENIV'+IntToStr(i))+'  ';
    end;
    CodeCol:=codeCol+' AND QBU_VALAXECU'+IntToStr(NivMax)+'="'+TF.GetValue('QBR_VALEURAXE')+'" ';
    codeTitre:=codeTitre+TabCodeAxe[NivMax+1]+' :'+TF.GetValue('QBR_VALEURAXE')+'  ';

    AglLanceFiche('Q','QUFSBPVUECLTDET','','','SESSION='+codeSession+';WHERE='+codeCol+';TITRE='+codeTitre+
                      ';COEFF='+CoeffRetenu);
  end
  else if TControl( Sender ).name = 'BTNCALREP' then
  begin
    dateDebFinSesssion(codeSession,DateDebC,DateFinC);

    case ValeurAffiche of
      1 : begin
            Prevu := TF.GetValue('QBR_OP1');
            Histo := TF.GetValue('QBR_REF1');
          end;
      2 : begin
            Prevu := TF.GetValue('QBR_QTEC');
            Histo := TF.GetValue('QBR_QTEREF');
          end;
      3 : begin
            Prevu := TF.GetValue('QBR_OP2');
            Histo := TF.GetValue('QBR_REF2');
          end;
      4 : begin
            Prevu := TF.GetValue('QBR_OP3');
            Histo := TF.GetValue('QBR_REF3');
          end;
      5 : begin
            Prevu := TF.GetValue('QBR_OP4');
            Histo := TF.GetValue('QBR_REF4');
          end;
      6 : begin
            Prevu := TF.GetValue('QBR_OP5');
            Histo := TF.GetValue('QBR_REF5');
          end;
      7 : begin
            Prevu := TF.GetValue('QBR_OP6');
            Histo := TF.GetValue('QBR_REF6');
          end;
    end;

    if TV.Selected.Level > 0 then
    begin
      for i := 1 to (TV.Selected.Level-1) do
      begin
        if ((TF.GetValue('QBR_VALAXENIV'+IntToStr(i)) = '') OR (TF.GetValue('QBR_VALAXENIV'+IntToStr(i)) = '[---]'))
        then Axes := Axes + 'VALAXENIV'+IntToStr(i)+'=[---];'
        else Axes := Axes + 'VALAXENIV'+IntToStr(i)+'='+TF.GetValue('QBR_VALAXENIV'+IntToStr(i))+';';
      end;
      if Tob(TV.Selected.data).GetValue('SOUSNIVEAU')<> 0 then
      begin
        if ((TF.GetValue('QBR_VALEURAXE') = '') OR (TF.GetValue('QBR_VALEURAXE') = '[---]'))
        then Axes :=  Axes + 'VALAXENIV'+IntToStr(TV.Selected.Level)+'=[---];'
        else Axes :=  Axes + 'VALAXENIV'+IntToStr(TV.Selected.Level)+'='+TF.GetValue('QBR_VALEURAXE')+';';
      end;
    end;

    aglLanceFiche('Q', 'QUFVBPCALENDJOURM', '', '','SESSION='+codeSession+';'
                  +Axes+';MODEFICHE=CALENDCONSULT;PREVU='+FloatToStr(Prevu)+';HISTO='+FloatToStr(Histo));
  end;
end;

procedure TOF_QUFSLARBRE.DoSetNavigate(Sender: TObject);
var niv,nivPrev,i,j,k,NivArt:integer;
    codeArt:hString;
    Enr:TEnreg;
begin
  niv:=nivMax;
  if SessionDelai(codeSession) then niv:=nivMax+1;

  if not sousNivOk then
  begin
    codeaxeT:=TF.GetValue('QBR_CODEAXE');
    numnoeudT:=IntToStr(BPIncrementenumNoeud(codeSession));
    numnoeudpereT:=TF.GetValue('QBR_NUMNOEUDPERE');
    niveauT:=TF.GetValue('QBR_NIVEAU');
    ValAxeNiv1T:=TF.GetValue('QBR_VALAXENIV1');
    ValAxeNiv2T:=TF.GetValue('QBR_VALAXENIV2');
    ValAxeNiv3T:=TF.GetValue('QBR_VALAXENIV3');
    ValAxeNiv4T:=TF.GetValue('QBR_VALAXENIV4');
    ValAxeNiv5T:=TF.GetValue('QBR_VALAXENIV5');
    ValAxeNiv6T:=TF.GetValue('QBR_VALAXENIV6');
    ValAxeNiv7T:=TF.GetValue('QBR_VALAXENIV7');
    ValAxeNiv8T:=TF.GetValue('QBR_VALAXENIV8');
    ValAxeNiv9T:=TF.GetValue('QBR_VALAXENIV9');
    ValAxeNiv10T:=TF.GetValue('QBR_VALAXENIV10');

    //change à chaque fois que l'on clique sur un noeud
    THLabel(GetControl( 'LBLCODEAXE' )).Caption:=codeaxeT;
    THLabel(GetControl( 'TQBR_VALEURAXE' )).Caption:=codeaxeT;
    if not BPOkOrli
    then
    begin
      THLabel(GetControl( 'LBLCODEAXE' )).Caption:=DonneLibelleCAxe('','',codeaxeT);
      THLabel(GetControl( 'TQBR_VALEURAXE' )).Caption:=DonneLibelleCAxe('','',codeaxeT);
    end;

    sousNivOk:=false;
  end;

  nivPrev:=ChercheNivMaxSession(codeSession);

  if VALEURI(TF.GetValue('QBR_NIVEAU'))=1 then TToolBarButton97(GetControl( 'BTNEVOLUTION' )).Enabled := false
  else TToolBarButton97(GetControl( 'BTNEVOLUTION' )).Enabled := true;

  if VALEURI(TF.GetValue('QBR_NIVEAU'))=NivMax then TToolBarButton97(GetControl( 'BTNCSUBNIVAUTO' )).Enabled := true
  else TToolBarButton97(GetControl( 'BTNCSUBNIVAUTO' )).Enabled := false;

  if VALEURI(TF.GetValue('QBR_NIVEAU'))=NivMax+1 then TToolBarButton97(GetControl( 'BTNMSUBNIVAUTO' )).Enabled := false
  else TToolBarButton97(GetControl( 'BTNMSUBNIVAUTO' )).Enabled := true;

  if TF.FirstNode
  then TToolBarButton97(GetControl( 'BINSERT' )).enabled := false
  else TToolBarButton97(GetControl( 'BINSERT' )).enabled := true;

  if consult='CONSULTATION' then
  begin
    TToolBarButton97(GetControl( 'BTNEVOLUTION' )).Enabled := false;
    TToolBarButton97(GetControl( 'BTNCSUBNIVAUTO' )).Enabled := false;
    TToolBarButton97(GetControl( 'BTNMSUBNIVAUTO' )).Enabled := false;
    TToolBarButton97(GetControl( 'BSOUSNIVEAU' )).Enabled := false;
    TToolBarButton97(GetControl( 'BDUPLIQUER' )).Enabled := false;
    TToolBarButton97(GetControl( 'BINSERT' )).Enabled := false;
  end
  else TToolBarButton97(GetControl( 'BSOUSNIVEAU' )).Enabled := true;

  if ContextBP=3 then
  begin
    TToolBarButton97(GetControl( 'BTNCSUBNIVAUTO' )).visible := true;
    TToolBarButton97(GetControl( 'BTNMSUBNIVAUTO' )).visible := true
  end
  else
  begin
    TToolBarButton97(GetControl( 'BTNCSUBNIVAUTO' )).visible := true;
    TToolBarButton97(GetControl( 'BTNMSUBNIVAUTO' )).visible := true
  end;

  if (not OkSessionInitPrev) then
  begin
    if VALEURI(TF.GetValue('QBR_NIVEAU'))>=niv
    then TToolBarButton97(GetControl( 'BSOUSNIVEAU' )).Enabled := false
  end
  else
  begin
    if VALEURI(TF.GetValue('QBR_NIVEAU'))>=nivPrev
    then TToolBarButton97(GetControl( 'BSOUSNIVEAU' )).Enabled := false;
  end;

  for i:=1 to TF.LaGrid.colcount do
  begin
    if TF.LaGrid.ColNames[i]='QBR_VALEURAXE' then
    begin
      TF.LaGrid.Cells[i,0]:=codeaxeT;
      if not BPOkOrli then TF.LaGrid.Cells[i,0]:=DonneLibelleCAxe('','',codeaxeT);
    end;
    if TF.LaGrid.ColNames[i]='QBR_LIBVALAXE'
    then
    begin
      TF.LaGrid.Cells[i,0]:='Libellé '+codeaxeT;
      if not BPOkOrli then TF.LaGrid.Cells[i,0]:=DonneLibelleCAxe('','',codeaxeT);
    end;

    { Prise en charge des ETP  }
    if ((OldCodeAxe<>codeaxeT) OR (NodeDelete=1)) then
    begin
      if ContextBp=3 then
      begin
        if codeaxeT='DELAI' then break;
        if TabValues[1]='ETP' then
        begin
          if TF.LaGrid.ColNames[i]='QBR_REF1' then
          begin
            For j:=1 to TF.LaGrid.RowCount-1 do
            TF.LaGrid.Cells[i,j]:=FloatToStr(VALEUR(TF.LaGrid.Cells[i,j])/NbIntervalle);
          end;
          if TF.LaGrid.ColNames[i]='QBR_OP1' then
          begin
            For j:=1 to TF.LaGrid.RowCount-1 do
            TF.LaGrid.Cells[i,j]:=FloatToStr(VALEUR(TF.LaGrid.Cells[i,j])/NbIntervalle);
          end;
        end;
        if TabValues[2]='ETP' then
        begin
          if TF.LaGrid.ColNames[i]='QBR_QTEREF' then
          begin
            For j:=1 to TF.LaGrid.RowCount-1 do
            TF.LaGrid.Cells[i,j]:=FloatToStr(VALEUR(TF.LaGrid.Cells[i,j])/NbIntervalle);
          end;
          if TF.LaGrid.ColNames[i]='QBR_QTEC' then
          begin
            For j:=1 to TF.LaGrid.RowCount-1 do
            TF.LaGrid.Cells[i,j]:=FloatToStr(VALEUR(TF.LaGrid.Cells[i,j])/NbIntervalle);
          end;
        end;
        for k:=3 to 7 do
        begin
          if TabValues[k]='ETP' then
          begin
            if TF.LaGrid.ColNames[i]='QBR_REF'+IntToStr(k-1) then
            begin
              For j:=1 to TF.LaGrid.RowCount-1 do
              TF.LaGrid.Cells[i,j]:=FloatToStr(VALEUR(TF.LaGrid.Cells[i,j])/NbIntervalle);
            end;
            if TF.LaGrid.ColNames[i]='QBR_OP'+IntToStr(k-1) then
            begin
              For j:=1 to TF.LaGrid.RowCount-1 do
              TF.LaGrid.Cells[i,j]:=FloatToStr(VALEUR(TF.LaGrid.Cells[i,j])/NbIntervalle);
            end;
          end
        end
      end;
    end
  end;
  OldCodeAxe:=codeaxeT;
  NodeDelete:=NodeDelete-1;

  if (not OkSessionObj) and (niveauT=IntToStr(nivMax))
  then
  begin
    THTabSheet(GetControl('TAB8')).tabVisible:=true;
    TToolBarButton97(GetControl( 'BTNDETCLT' )).visible:=true;
  end
  else
  begin
    THTabSheet(GetControl('TAB8')).tabVisible:=false;
    TToolBarButton97(GetControl( 'BTNDETCLT' )).visible:=false;
  end;

  if (OkSessionInitPrev)
  then
  begin
    THTabSheet(GetControl('TAB10')).tabVisible:=true;

    if (niveauT>IntToStr(nivMax))
    then THTabSheet(GetControl('TAB11')).tabVisible:=false
    else
    begin
      if DonneParamS(ps_BPCoeffPerCAQte)='2'
      then THTabSheet(GetControl('TAB11')).tabVisible:=false
      else THTabSheet(GetControl('TAB11')).tabVisible:=true;
    end;
  end;

  if (not OkSessionObj) and (niveauT=IntToStr(nivPrev))
  then
  begin
    if codeaxeT='ARTICLE'
    then codeArt:=TF.GetValue('QBR_VALEURAXE')
    else
    begin
      NivArt:=ChercheNivMax(codeSession)+1;
      codeArt:=TF.GetValue('QBR_VALAXENIV'+IntToStr(NivArt));
    end;
    if GestionBPTaille
    then
    begin
      //grille de taille de l'article
      if dm_trouveEnr(idm_tmpBPArticleTai,[codeArt],Enr)=0
      then SetControlText('GRILLETAILLE',Enr.ch(BPARTICLETAI_GRILLETAI));
      if (SessionDelai(codeSession)) and  ((SessionCalculParTaille(codeSession)) or (SessionEclateeParTaille(codeSession)))
      then THTabSheet(GetControl('TAB9')).tabVisible:=true;
    end;
  end
  else THTabSheet(GetControl('TAB9')).tabVisible:=false;

  if not BPOkOrli
  then
  Case ContextBP of
    0,1,2 : begin
            case ValeurAffiche of
             1 : THTabSheet(GetControl('TAB1')).tabVisible:=true;
             2 : THTabSheet(GetControl('TAB2')).tabVisible:=true;
             3 : THTabSheet(GetControl('TAB3')).tabVisible:=true;
             4 : THTabSheet(GetControl('TAB4')).tabVisible:=true;
             5 : THTabSheet(GetControl('TAB5')).tabVisible:=true;
             6 : THTabSheet(GetControl('TAB6')).tabVisible:=true;
             7 : THTabSheet(GetControl('TAB7')).tabVisible:=true;
            end;
          end;
  end;
end;

{ EVI / Correction bug redimensionnement colonnes si clic sur élément de l'arbre }
{ procedure TOF_QUFSLARBRE.Resize;
var
  largeur,i : integer;
begin
  largeur := ThGrid( GetControl( 'FListe' ) ).Width;
  largeur := Trunc( ( largeur ) / (TF.LaGrid.colcount-1) );
  for i:=1 to TF.LaGrid.colcount-1 do TF.LaGrid.ColWidths[i]:=largeur - 3;
end; }
procedure TOF_QUFSLARBRE.Resize;
var NBColonne,largeur,i,j : integer;
    TabIndiceCol:array of integer;
begin
  NbColonne:=0;
  largeur := ThGrid( GetControl( 'FListe' ) ).Width;
  //Nombre de colonnes > -1
  for i:=1 to TF.LaGrid.colcount-1 do
  begin
    if TF.LaGrid.ColWidths[i] > -1 then NbColonne := NbColonne + 1;
  end;
  SetLength(TabIndiceCol,NbColonne);
  //Renseigne le tableau avec les indices des colonnes supérieures à -1
  j:=0;
  for i:=1 to TF.LaGrid.colcount-1 do
  begin
    if TF.LaGrid.ColWidths[i] > -1 then
    begin
      TabIndiceCol[j]:=i;
      j:=j+1;
    end;
  end;
  if NbColonne <> 0 then
  begin
    largeur := Trunc( ( largeur ) / (NbColonne) );
    for i:=0 to NbColonne-1 do TF.LaGrid.ColWidths[TabIndiceCol[i]]:=largeur ;
  end;
end;


procedure TOF_QUFSLARBRE.PageChange(Sender: TObject);
var NameColumn:hString;
    i:integer;
    LgColValeur,LgColRef,LgColOP,LgColRefPrct,LgColOPPrct:integer;
    LgPrctVariation,LgColCache:integer;
    TabRef,TabOP,TabRefPrct,TabOPPRct,TabPrctVariation : array [0..7] of integer;
begin
  TabActive:=THPageControl2(GetControl('PAGE' )).ActivePage.Name;
  LgColValeur:=120;
  LgColRef:=120;
  LgColOP:=120;
  LgColRefPrct:=120;
  LgColOPPrct:=120;
  LgPrctVariation:=101;
  LgColCache:=-1;

  for i := 0 to 7 do                                // GEY changement 1 --> 8 par  0--> 7 sinon plantage
  begin
    if i=StrToInt(AnsiRightStr(TabActive,1)) then
    begin
      TabRef[i]:=LgColRef;
      TabOP[i]:=LgColOP;
      TabRefPrct[i]:=LgColRefPrct;
      TabOPPRct[i]:=LgColOPPrct;
      TabPrctVariation[i]:=LgPrctVariation;
    end
    else
    begin
      TabRef[i]:=LgColCache;
      TabOP[i]:=LgColCache;
      TabRefPrct[i]:=LgColCache;
      TabOPPRct[i]:=LgColCache;
      TabPrctVariation[i]:=LgColCache;
    end;
  end;

  if OkSessionObj then NameColumn:=TraduireMemoire('Objectif') else NameColumn:=TraduireMemoire('Prévu');

  for i:=1 to TF.LaGrid.colcount-1 do
  begin
    if TF.LaGrid.ColNames[i]='QBR_OP1' then TF.LaGrid.Cells[i,0]:=NameColumn;
    if TF.LaGrid.ColNames[i]='QBR_OP2' then TF.LaGrid.Cells[i,0]:=NameColumn;
    if TF.LaGrid.ColNames[i]='QBR_OP3' then TF.LaGrid.Cells[i,0]:=NameColumn;
    if TF.LaGrid.ColNames[i]='QBR_OP4' then TF.LaGrid.Cells[i,0]:=NameColumn;
    if TF.LaGrid.ColNames[i]='QBR_OP5' then TF.LaGrid.Cells[i,0]:=NameColumn;
    if TF.LaGrid.ColNames[i]='QBR_OP6' then TF.LaGrid.Cells[i,0]:=NameColumn;
    if TF.LaGrid.ColNames[i]='QBR_QTEC' then TF.LaGrid.Cells[i,0]:=NameColumn;

    if TF.LaGrid.ColNames[i]='QBR_VALEURAXE' then TF.LaGrid.ColWidths[i]:=LgColValeur;

    //TAB1
    if TF.LaGrid.ColNames[i]='QBR_REF1' then TF.LaGrid.ColWidths[i]:=TabRef[1];
    if TF.LaGrid.ColNames[i]='QBR_OP1' then TF.LaGrid.ColWidths[i]:=TabOP[1];
    if TF.LaGrid.ColNames[i]='QBR_REFPRCT1' then TF.LaGrid.ColWidths[i]:=TabRefPrct[1];
    if TF.LaGrid.ColNames[i]='QBR_OPPRCT1' then TF.LaGrid.ColWidths[i]:=TabOPPRct[1];
    if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION1' then TF.LaGrid.ColWidths[i]:=TabPrctVariation[1];
    //TAB2
    if TF.LaGrid.ColNames[i]='QBR_QTEREF' then TF.LaGrid.ColWidths[i]:=TabRef[2];
    if TF.LaGrid.ColNames[i]='QBR_QTEREFPRCT' then TF.LaGrid.ColWidths[i]:=TabOP[2];
    if TF.LaGrid.ColNames[i]='QBR_QTEC' then TF.LaGrid.ColWidths[i]:=TabRefPrct[2];
    if TF.LaGrid.ColNames[i]='QBR_QTECPRCT' then TF.LaGrid.ColWidths[i]:=TabOPPRct[2];
    if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATIONQ' then TF.LaGrid.ColWidths[i]:=TabPrctVariation[2];
    //TAB3
    if TF.LaGrid.ColNames[i]='QBR_REF2' then TF.LaGrid.ColWidths[i]:=TabRef[3];
    if TF.LaGrid.ColNames[i]='QBR_OP2' then TF.LaGrid.ColWidths[i]:=TabOP[3];
    if TF.LaGrid.ColNames[i]='QBR_REFPRCT2' then TF.LaGrid.ColWidths[i]:=TabRefPrct[3];
    if TF.LaGrid.ColNames[i]='QBR_OPPRCT2' then TF.LaGrid.ColWidths[i]:=TabOPPRct[3];
    if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION2' then TF.LaGrid.ColWidths[i]:=TabPrctVariation[3];
    //TAB4
    if TF.LaGrid.ColNames[i]='QBR_REF3' then TF.LaGrid.ColWidths[i]:=TabRef[4];
    if TF.LaGrid.ColNames[i]='QBR_OP3' then TF.LaGrid.ColWidths[i]:=TabOP[4];
    if TF.LaGrid.ColNames[i]='QBR_REFPRCT3' then TF.LaGrid.ColWidths[i]:=TabRefPrct[4];
    if TF.LaGrid.ColNames[i]='QBR_OPPRCT3' then TF.LaGrid.ColWidths[i]:=TabOPPRct[4];
    if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION3' then TF.LaGrid.ColWidths[i]:=TabPrctVariation[4];
    //TAB5
    if TF.LaGrid.ColNames[i]='QBR_REF4' then TF.LaGrid.ColWidths[i]:=TabRef[5];
    if TF.LaGrid.ColNames[i]='QBR_OP4' then TF.LaGrid.ColWidths[i]:=TabOP[5];
    if TF.LaGrid.ColNames[i]='QBR_REFPRCT4' then TF.LaGrid.ColWidths[i]:=TabRefPrct[5];
    if TF.LaGrid.ColNames[i]='QBR_OPPRCT4' then TF.LaGrid.ColWidths[i]:=TabOPPRct[5];
    if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION4' then TF.LaGrid.ColWidths[i]:=TabPrctVariation[5];
    //TAB6
    if TF.LaGrid.ColNames[i]='QBR_REF5' then TF.LaGrid.ColWidths[i]:=TabRef[6];
    if TF.LaGrid.ColNames[i]='QBR_OP5' then TF.LaGrid.ColWidths[i]:=TabOP[6];
    if TF.LaGrid.ColNames[i]='QBR_REFPRCT5' then TF.LaGrid.ColWidths[i]:=TabRefPrct[6];
    if TF.LaGrid.ColNames[i]='QBR_OPPRCT5' then TF.LaGrid.ColWidths[i]:=TabOPPRct[6];
    if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION5' then TF.LaGrid.ColWidths[i]:=TabPrctVariation[6];
    //TAB7
    if TF.LaGrid.ColNames[i]='QBR_REF6' then TF.LaGrid.ColWidths[i]:=TabRef[7];
    if TF.LaGrid.ColNames[i]='QBR_OP6' then TF.LaGrid.ColWidths[i]:=TabOP[7];
    if TF.LaGrid.ColNames[i]='QBR_REFPRCT6' then TF.LaGrid.ColWidths[i]:=TabRefPrct[7];
    if TF.LaGrid.ColNames[i]='QBR_OPPRCT6' then TF.LaGrid.ColWidths[i]:=TabOPPRct[7];
    if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION6' then TF.LaGrid.ColWidths[i]:=TabPrctVariation[7];

    //Coeff
    if TF.LaGrid.ColNames[i]='QBR_HISTO' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_REALISE' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_PREVU' then TF.LaGrid.ColWidths[i]:=LgColCache;

    if TF.LaGrid.ColNames[i]='QBR_COEFFCALCUL' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_COEFFRETENU' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_QTERETENUE' then TF.LaGrid.ColWidths[i]:=LgColCache;

    if TF.LaGrid.ColNames[i]='QBR_HISTOCA' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_PREVUCA' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_REALISECA' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_CARETENU' then TF.LaGrid.ColWidths[i]:=LgColCache;

    if TF.LaGrid.ColNames[i]='QBR_REALISE' then TF.LaGrid.Cells[i,0]:='Réalisé';
    if TF.LaGrid.ColNames[i]='QBR_PREVU' then TF.LaGrid.Cells[i,0]:='Prévu';
  end;

  //Coeff
  if TabActive='TAB10' then
  for i:=1 to TF.LaGrid.colcount do
  begin
    if DonneParamS(ps_BPCoeffPerCAQte)='2' then
    begin
      if TF.LaGrid.ColNames[i]='QBR_HISTO'  then TF.LaGrid.ColWidths[i]:=LgColOP;
      if TF.LaGrid.ColNames[i]='QBR_PREVU'  then TF.LaGrid.ColWidths[i]:=LgColOP;
      if TF.LaGrid.ColNames[i]='QBR_REALISE' then TF.LaGrid.ColWidths[i]:=LgColOP;
      if TF.LaGrid.ColNames[i]='QBR_HISTOCA' then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PREVUCA' then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REALISECA' then TF.LaGrid.ColWidths[i]:=LgColCache;
    end
    else
    begin
      if TF.LaGrid.ColNames[i]='QBR_HISTO' then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PREVU' then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REALISE' then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_HISTOCA' then TF.LaGrid.ColWidths[i]:=LgColOP;
      if TF.LaGrid.ColNames[i]='QBR_PREVUCA' then TF.LaGrid.ColWidths[i]:=LgColOP;
      if TF.LaGrid.ColNames[i]='QBR_REALISECA' then TF.LaGrid.ColWidths[i]:=LgColOP;
    end;
    if TF.LaGrid.ColNames[i]='QBR_COEFFCALCUL' then TF.LaGrid.ColWidths[i]:=LgColOP;
    if TF.LaGrid.ColNames[i]='QBR_COEFFRETENU' then TF.LaGrid.ColWidths[i]:=LgColOP;
    if TF.LaGrid.ColNames[i]='QBR_QTERETENUE'  then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_CARETENU' then TF.LaGrid.ColWidths[i]:=LgColCache;
  end;

  //Quantite prev
  if TabActive='TAB10' then
  for i:=1 to TF.LaGrid.colcount do
  begin
    if TF.LaGrid.ColNames[i]='QBR_HISTO'  then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_REALISE' then TF.LaGrid.ColWidths[i]:=LgColValeur;
    if TF.LaGrid.ColNames[i]='QBR_PREVU'  then TF.LaGrid.ColWidths[i]:=LgColValeur;
    if TF.LaGrid.ColNames[i]='QBR_HISTOCA' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_PREVUCA' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_REALISECA' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_COEFFCALCUL' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_QTERETENUE' then TF.LaGrid.ColWidths[i]:=LgColValeur;
    if TF.LaGrid.ColNames[i]='QBR_CARETENU' then TF.LaGrid.ColWidths[i]:=LgColCache;

    if (VALEUR(niveauT)>nivMax) then
    begin
      if TF.LaGrid.ColNames[i]='QBR_COEFFRETENU' then TF.LaGrid.ColWidths[i]:=LgColValeur;
    end
    else
    begin
      if TF.LaGrid.ColNames[i]='QBR_COEFFRETENU' then TF.LaGrid.ColWidths[i]:=LgColCache;
    end;
  end;

  //CA prev
  if TabActive='TAB11' then
  for i:=1 to TF.LaGrid.colcount do
  begin
    //Coeff
    if TF.LaGrid.ColNames[i]='QBR_HISTO' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_REALISE' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_PREVU' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_HISTOCA' then TF.LaGrid.ColWidths[i]:=LgColValeur;
    if TF.LaGrid.ColNames[i]='QBR_PREVUCA' then TF.LaGrid.ColWidths[i]:=LgColValeur;
    if TF.LaGrid.ColNames[i]='QBR_REALISECA' then TF.LaGrid.ColWidths[i]:=LgColValeur;
    if TF.LaGrid.ColNames[i]='QBR_COEFFCALCUL' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_COEFFRETENU' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_QTERETENUE' then TF.LaGrid.ColWidths[i]:=LgColCache;
    if TF.LaGrid.ColNames[i]='QBR_CARETENU' then TF.LaGrid.ColWidths[i]:=LgColValeur;
  end;

  Resize;
end;

procedure TOF_QUFSLARBRE.TreeViewPopUp(Sender: TObject);
begin
  TMenuItem(GetControl('CALENDVIEW')).OnClick := CalendView_OnClick;
  TMenuItem(GetControl('CALENDALLHISTO')).OnClick := CalendAllHisto_OnClick;
  TMenuItem(GetControl('CALENDALLTYPE')).OnClick := CalendAllType_OnClick;
  TMenuItem(GetControl('CALENDALLDELETE')).OnClick := CalendAllDelete_OnClick;
  TMenuItem(GetControl('CALENDHISTO')).OnClick := CalendHisto_OnClick;
  TMenuItem(GetControl('CALENDTYPE')).OnClick := CalendType_OnClick;
  TMenuItem(GetControl('CALENDDELETE')).OnClick := CalendDelete_OnClick;
  TMenuItem(GetControl('IMPORTDETCALEND')).OnClick := ImportDetCalend_OnClick;

  TMenuItem(GetControl('MSOUSNIVOAUTO')).OnClick := ModifSousNiveauAuto_OnClick;
  TMenuItem(GetControl('CSOUSNIVOAUTO')).OnClick := CreateSousNiveauAuto_OnClick;
  TMenuItem(GetControl('AXEEVOLUTION')).OnClick := ProcBtnEvolution;
  TMenuItem(GetControl('NEWSOUSNIVO')).OnClick := ProcBSousNiveau;
  TMenuItem(GetControl('DUPLIKNODE')).OnClick := ProcBDupliquer;

  TMenuItem(GetControl('REFRESH')).OnClick := Refresh_OnClick;
  TMenuItem(GetControl('EXPAND')).OnClick := Expand_OnClick;

  if ContextBP in [0,1,2] then
  begin
    TMenuItem(GetControl('CSOUSNIVOAUTO')).visible := false;
    TMenuItem(GetControl('MSOUSNIVOAUTO')).visible := false;
    if BPInitialise <> '0' then
    begin
      if Tob(TV.Selected.data).GetValue('QBG_NUMNOEUD') = 0 then
      begin
        TMenuItem(GetControl('CALENDALLHISTO')).Visible := true;
        TMenuItem(GetControl('CALENDALLTYPE')).Visible := true;
        TMenuItem(GetControl('CALENDALLDELETE')).visible := true;
        TMenuItem(GetControl('CALENDHISTO')).visible := false;
        TMenuItem(GetControl('CALENDTYPE')).visible :=  false;
        TMenuItem(GetControl('CALENDVIEW')).visible := false;
        TMenuItem(GetControl('CALENDDELETE')).visible := false;
        TMenuItem(GetControl('IMPORTDETCALEND')).visible := false;
        TMenuItem(GetControl('CALENDLINE')).Visible := true;
      end
      else
      begin
        if Tob(TV.Selected.data).GetValue('QBG_NIVEAU') = NivMax+1 then
        begin
          TMenuItem(GetControl('CALENDALLHISTO')).Visible := false;
          TMenuItem(GetControl('CALENDALLTYPE')).Visible := false;
          TMenuItem(GetControl('CALENDALLDELETE')).visible := false;
          TMenuItem(GetControl('CALENDHISTO')).visible := false;
          TMenuItem(GetControl('CALENDTYPE')).visible :=  false;
          TMenuItem(GetControl('CALENDVIEW')).visible := false;
          TMenuItem(GetControl('IMPORTDETCALEND')).visible := false;
          TMenuItem(GetControl('CALENDDELETE')).visible := false;
          TMenuItem(GetControl('CALENDLINE')).Visible := false;
        end
        else
        begin

          if TOB(TV.Selected.Data).FieldExists('CALEND') then
          begin
            if TOB(TV.Selected.Data).GetValue('CALEND') = 'X' then
            begin
              TMenuItem(GetControl('CALENDALLHISTO')).Visible := false;
              TMenuItem(GetControl('CALENDALLTYPE')).Visible := false;
              TMenuItem(GetControl('CALENDALLDELETE')).visible := false;
              TMenuItem(GetControl('CALENDHISTO')).visible := false;
              TMenuItem(GetControl('CALENDTYPE')).visible :=  false;
              TMenuItem(GetControl('CALENDVIEW')).visible := true;
              TMenuItem(GetControl('CALENDDELETE')).visible := true;
              TMenuItem(GetControl('IMPORTDETCALEND')).visible := true;
              TMenuItem(GetControl('CALENDLINE')).Visible := true;
            end
            else
            begin
              TMenuItem(GetControl('CALENDALLHISTO')).Visible := false;
              TMenuItem(GetControl('CALENDALLTYPE')).Visible := false;
              TMenuItem(GetControl('CALENDALLDELETE')).visible := false;
              TMenuItem(GetControl('CALENDHISTO')).visible := true;
              TMenuItem(GetControl('CALENDTYPE')).visible := true;
              TMenuItem(GetControl('CALENDVIEW')).visible := false;
              TMenuItem(GetControl('CALENDDELETE')).visible := false;
              TMenuItem(GetControl('IMPORTDETCALEND')).visible := false;
              TMenuItem(GetControl('CALENDLINE')).Visible := true;
            end;

          end;
        end;
      end;
    end
    else
    begin
      //Loi d'Eclatement
      TMenuItem(GetControl('CALENDALLHISTO')).Visible := false;
      TMenuItem(GetControl('CALENDALLTYPE')).Visible := false;
      TMenuItem(GetControl('CALENDALLDELETE')).visible := false;
      TMenuItem(GetControl('CALENDHISTO')).visible := false;
      TMenuItem(GetControl('CALENDTYPE')).visible :=  false;
      TMenuItem(GetControl('CALENDVIEW')).visible := false;
      TMenuItem(GetControl('IMPORTDETCALEND')).visible := false;
      TMenuItem(GetControl('CALENDDELETE')).visible := false;
      TMenuItem(GetControl('CALENDLINE')).Visible := false;
    end;
  end
  else
  begin
    //Paie
    TMenuItem(GetControl('CALENDALLHISTO')).Visible := false;
    TMenuItem(GetControl('CALENDALLTYPE')).Visible := false;
    TMenuItem(GetControl('CALENDALLDELETE')).visible := false;
    TMenuItem(GetControl('CALENDHISTO')).visible := false;
    TMenuItem(GetControl('CALENDTYPE')).visible :=  false;
    TMenuItem(GetControl('CALENDVIEW')).visible := false;
    TMenuItem(GetControl('CALENDDELETE')).visible := false;
    TMenuItem(GetControl('IMPORTDETCALEND')).visible := false;
    TMenuItem(GetControl('CALENDLINE')).Visible := false;
  end;

  if Tob(TV.Selected.data).GetValue('QBG_NIVEAU') = 0 then
  begin
    TMenuItem(GetControl('DUPLIKNODE')).visible := false;
    TMenuItem(GetControl('MSOUSNIVOAUTO')).visible := false;
    TMenuItem(GetControl('NODELINE')).visible := false;
  end
  else
  begin
    TMenuItem(GetControl('DUPLIKNODE')).visible := true;
    TMenuItem(GetControl('MSOUSNIVOAUTO')).visible := true;
    TMenuItem(GetControl('NODELINE')).visible := true;
  end;

  if Tob(TV.Selected.data).GetValue('QBG_NIVEAU') > 1
  then TMenuItem(GetControl('AXEEVOLUTION')).visible := true
  else TMenuItem(GetControl('AXEEVOLUTION')).visible := false;

  if Tob(TV.Selected.data).GetValue('QBG_NIVEAU') = NivMax
  then TMenuItem(GetControl('CSOUSNIVOAUTO')).visible := true
  else TMenuItem(GetControl('CSOUSNIVOAUTO')).visible := false;

  if Tob(TV.Selected.data).GetValue('QBG_NIVEAU') = NivMax+1 then
  begin
    TMenuItem(GetControl('NEWSOUSNIVO')).visible := false;
    TMenuItem(GetControl('MSOUSNIVOAUTO')).visible := false;
  end
  else TMenuItem(GetControl('NEWSOUSNIVO')).visible := true;

  if consult='CONSULTATION' then
  begin
    TMenuItem(GetControl('CALENDVIEW')).visible := false;
    TMenuItem(GetControl('CALENDALLHISTO')).visible := false;
    TMenuItem(GetControl('CALENDALLTYPE')).visible := false;
    TMenuItem(GetControl('CALENDALLDELETE')).visible := false;
    TMenuItem(GetControl('CALENDHISTO')).visible := false;
    TMenuItem(GetControl('CALENDTYPE')).visible := false;
    TMenuItem(GetControl('CALENDDELETE')).visible := false;
    TMenuItem(GetControl('IMPORTDETCALEND')).visible := false;
    TMenuItem(GetControl('CSOUSNIVOAUTO')).visible := false;
    TMenuItem(GetControl('MSOUSNIVOAUTO')).visible := false;
    TMenuItem(GetControl('AXEEVOLUTION')).visible := false;
    TMenuItem(GetControl('NEWSOUSNIVO')).visible := false;
    TMenuItem(GetControl('DUPLIKNODE')).visible := false;
  end;

  if Tob(TV.Selected.data).GetValue('QBG_NIVEAU') = NivMax+1 then TMenuItem(GetControl('EXPAND')).Visible := false
  else
  begin
    TMenuItem(GetControl('EXPAND')).Visible := true;
    if TV.Selected.Expanded = true then TMenuItem(GetControl('EXPAND')).Caption := 'Replier le noeud'
    else TMenuItem(GetControl('EXPAND')).Caption := 'Déployer le noeud';
  end;
  TMenuItem(GetControl('EXPANDALL')).OnClick := ExpandAll_OnClick;
  TMenuItem(GetControl('COLLAPSEALL')).OnClick := CollapseAll_OnClick;
  { EVI / Garder le focus sur le noeud sélectionné par clic droit }
  TV.Selected.Selected := true;
end;

procedure TOF_QUFSLARBRE.CalendDelete_OnClick(Sender: TObject);
var Axes:hString;
level,i,DelOK:integer;
  SelectedNode :TTreeNode;
begin
  SelectedNode := TV.Selected;
  Level := TV.Selected.Level;
  if HShowmessage('1;Suppression;Vous allez supprimer le calendrier associé à votre sélection.#13#10 Voulez-vous continuer ?;Q;YN;N;N;', '', '') = MrYes
  then
  begin
    blocageMenu(TFSaisieList(Ecran));
    TF.State := dsEdit;
    SelectedNode.Selected:=true;
    if Level > 0 then
    begin
      for i := 1 to (Level-1) do
      begin
        if ((TF.GetValue('QBR_VALAXENIV'+IntToStr(i)) = '') OR (TF.GetValue('QBR_VALAXENIV'+IntToStr(i)) = '[---]'))
        then Axes := Axes + 'AND QBE_VALAXENIV'+IntToStr(i)+'="[---]"'
        else Axes := Axes + 'AND QBE_VALAXENIV'+IntToStr(i)+'="'+TF.GetValue('QBR_VALAXENIV'+IntToStr(i))+'"';
      end;
      if Tob(TV.Selected.data).GetValue('SOUSNIVEAU')<> 0 then
      begin
        if ((TF.GetValue('QBR_VALEURAXE') = '') OR (TF.GetValue('QBR_VALEURAXE') = '[---]'))
        then Axes :=  Axes + 'AND QBE_VALAXENIV'+IntToStr(Level)+'="[---]"'
        else Axes :=  Axes + 'AND QBE_VALAXENIV'+IntToStr(Level)+'="'+TF.GetValue('QBR_VALEURAXE')+'"';
      end;
    end;
    DelOK := ExecuteSQL( 'DELETE FROM QBPDETCALENDREP WHERE QBE_CALENDREP="'+codeSession+'"'+Axes  );
    TF.Post;
    if DelOK > 0 then HShowmessage('1;Suppression;Le calendrier a été supprimé.;X;O;', '', '')
    else HShowmessage('1;Suppression;Aucun enregistrement trouvé.;E;O;', '', '');
    deblocageMenu(TFSaisieList(Ecran));
  end;
end;

procedure TOF_QUFSLARBRE.CalendAllDelete_OnClick(Sender: TObject);
var DelOK,i:integer;
SelectedNode:TTreeNode;
begin
  if HShowmessage('1;Suppression;Vous allez supprimer tous les calendriers associés à cette session.#13#10 Etes-vous sûr ?;Q;YN;N;N;', '', '') = MrYes
  then
  begin
    SelectedNode:=TV.Selected.getFirstChild;
    blocageMenu(TFSaisieList(Ecran));

    DelOK := ExecuteSQL( 'DELETE FROM QBPDETCALENDREP WHERE QBE_CALENDREP="'+codeSession+'"');
    TF.First;
    if DelOK > 0 then
    begin
      for i:=0 to TOB(TV.Selected.Data).Detail.Count-1 do
      begin
        TF.State := dsEdit;
        TF.Post;
        SelectedNode.Collapse(false);
        SelectedNode:=SelectedNode.getNextSibling;
        TF.Next;
      end;
    end;

    if DelOK > 0 then HShowmessage('1;Suppression;Tous les calendriers ont été supprimés.;X;O;', '', '')
    else HShowmessage('1;Suppression;Aucun enregistrement trouvé.;E;O;', '', '');
    deblocageMenu(TFSaisieList(Ecran));
  end;
end;

procedure  TOF_QUFSLARBRE.CalendView_OnClick(Sender: TObject);
begin
 DoClick(  TToolBarButton97(GetControl( 'BTNCALREP' )));
end;

procedure TOF_QUFSLARBRE.ImportDetCalend_OnClick(Sender: TObject);
var Calendrier:hString;
    TCalend : tob;
    i:integer;

  function LoadTobFromSql(Sql: hString; t:Tob; lAppend: boolean = false): boolean;
  var Q: tQuery;
  begin
    Result := false;
    Q := OpenSql(Sql, True);
    try
      if not Q.Eof then
      begin
        t.loadDetailDB('', '', '', Q, lAppend);
        Result := true;
      end;
    finally
      Ferme(Q);
    end;
  end;

begin
  Calendrier:=AglLanceFiche('Q','QUFMBPCALENDREP','','',codeSession);
  if ((Calendrier<>'') AND (Calendrier<>'CANCEL')) then
  begin
    TCalend := Tob.Create( 'Virtuelle ',nil,-1 );
    LoadTobFromSql('SELECT QBE_DATEDEBSEM FROM QBPDETCALENDREP WHERE QBE_VALAXENIV1 IS NULL AND QBE_CALENDREP="'+ Calendrier+'"',TCalend);
    for i := 0 to TCalend.Detail.Count - 1 do
    ExecuteSQL( 'UPDATE QBPDETCALENDREP SET QBE_PRCTJOUR1=(SELECT QBE_PRCTJOUR1 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"),'+
                        ' QBE_PRCTJOUR2=(SELECT QBE_PRCTJOUR2 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"),'+
                        ' QBE_PRCTJOUR3=(SELECT QBE_PRCTJOUR3 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"),'+
                        ' QBE_PRCTJOUR4=(SELECT QBE_PRCTJOUR4 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"),'+
                        ' QBE_PRCTJOUR5=(SELECT QBE_PRCTJOUR5 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"),'+
                        ' QBE_PRCTJOUR6=(SELECT QBE_PRCTJOUR6 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"),'+
                        ' QBE_PRCTJOUR7=(SELECT QBE_PRCTJOUR7 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"),'+
                        ' QBE_COMMENTAIRE1=(SELECT QBE_COMMENTAIRE1 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"),'+
                        ' QBE_COMMENTAIRE2=(SELECT QBE_COMMENTAIRE2 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"),'+
                        ' QBE_COMMENTAIRE3=(SELECT QBE_COMMENTAIRE3 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"),'+
                        ' QBE_COMMENTAIRE4=(SELECT QBE_COMMENTAIRE4 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"),'+
                        ' QBE_COMMENTAIRE5=(SELECT QBE_COMMENTAIRE5 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"),'+
                        ' QBE_COMMENTAIRE6=(SELECT QBE_COMMENTAIRE6 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"),'+
                        ' QBE_COMMENTAIRE7=(SELECT QBE_COMMENTAIRE7 FROM QBPDETCALENDREP OLDCALEND'+
                        ' WHERE QBPDETCALENDREP.QBE_DATEDEBSEM=OLDCALEND.QBE_DATEDEBSEM'+
                        ' AND OLDCALEND.QBE_VALAXENIV1 IS NULL AND OLDCALEND.QBE_CALENDREP="'+Calendrier+
                        '" AND OLDCALEND.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'")'+
                        ' FROM QBPDETCALENDREP WHERE QBPDETCALENDREP.QBE_CALENDREP="'+CodeSession+'" AND '+
                        ' QBPDETCALENDREP.QBE_DATEDEBSEM="'+USDateTime(TCalend.Detail[i].GetValue('QBE_DATEDEBSEM'))+'"');

    HShowmessage('1;'+TraduireMemoire('Mise à jour')+';'+IntToStr(TCalend.Detail.Count*7)+TraduireMemoire(' lignes du calendrier mise à jour.')+';E;O;O;O;', '', '');
    TCalend.Free;
  end;
end;

procedure TOF_QUFSLARBRE.CalendAllType_OnClick(Sender: TObject);
var i,j : integer;
Axes,CalendValid,Etab,PrevHisto,Calendrier,ValidationAuto: hString;
OKCalend:boolean;
SelectedNode : TTreeNode;
DateDebC,DateFinC:TDateTime;
Prevu,Histo:double;
Q:TQuery;
ReqSQL,ComaSQL : string;
begin
  Axes := '';
  Prevu:=0;
  Histo:=0;
  OKCalend := true;
  SelectedNode := TV.Selected.GetFirstChild;

  { EVI / Test des sous-niveaux }
  if existeSQL( 'SELECT ##TOP 1## QBE_CALENDREP FROM QBPDETCALENDREP WHERE QBE_CALENDREP="'+codeSession+'"'  ) then
  begin
    HShowmessage('1;Initialisation impossible;Il existe au moins un calendrier défini pour un des sous-niveaux de votre sélection.;W;O;', '', '');
    OKCalend := false;
  end;

  if OKCalend = true then
  begin
    if TOB(TV.Selected.Data).Detail.Count > 1 then
    begin
      if HShowmessage('1;Initialisation;'+IntToStr(TOB(TV.Selected.Data).Detail.Count)+' calendriers vont '+
                      'être générés.#13#10 Etes-vous sûr ?;Q;YN;N;N;', '', '') = Mrno
      then OKCalend := false;
    end
    else
    begin
      if HShowmessage('1;Initialisation;Un calendrier va être généré.#13#10 Etes-vous sûr ?;Q;YN;N;N;', '', '') = Mrno
      then OKCalend := false;
    end;
  end;

  if OKCalend = true then
  begin
    blocageMenu(TFSaisieList(Ecran));
    TF.First;

    Calendrier:=AglLanceFiche('Q','QUFMBPCALENDREP','','','CALENDTYPE');
    if (Calendrier <> 'CANCEL') AND (Calendrier <> '') then
    begin
      dateDebFinSesssion(codeSession,DateDebC,DateFinC);
      { EVI / Remplissage de la table BPCUBETMP }
      InitMoveProgressForm(nil,TraduireMemoire('Génération en cours'),TraduireMemoire('Veuillez patienter'), 1 + NbDeMaille(codesession,BPInitialise,0,0), True, True);
      RemplitTableQBPCubeTmpPgi(codeSession,'','',DateDebC,DateFinC,0,0);

      if not MoveCurProgressForm(TraduireMemoire('Génération...')) then
      begin
        FiniMoveProgressForm;
        TF.Cancel;
        exit;
      end;

      FiniMoveProgressForm;

      if HShowmessage('1;Calendriers;Voulez-vous que tous les calendriers soient validés automatiquement ?;Q;YN;N;N;', '', '') = MrYes
      then ValidationAuto:='TRUE' else ValidationAuto:='FALSE';

      case ValeurAffiche of
        1 : PrevHisto := 'QBR_OP1,QBR_REF1';
        2 : PrevHisto := 'QBR_QTEC,QBR_QTEREF';
        3 : PrevHisto := 'QBR_OP2,QBR_REF2';
        4 : PrevHisto := 'QBR_OP3,QBR_REF3';
        5 : PrevHisto := 'QBR_OP4,QBR_REF4';
        6 : PrevHisto := 'QBR_OP5,QBR_REF5';
        7 : PrevHisto := 'QBR_OP6,QBR_REF6';
      end;

      if ValidationAuto='TRUE' then
      InitMoveProgressForm(nil,TraduireMemoire('Enregistrement des calendriers'),TraduireMemoire('Veuillez patienter'), TOB(TV.Selected.Data).Detail.Count, False, True);

      for i:=0 to TOB(TV.Selected.Data).Detail.Count-1 do
      begin
        TF.State := dsEdit;
        Etab:=TOB(TV.Selected.Data).Detail[i].GetString('QBG_VALEURAXE');

        if ((Etab = '') OR (Etab = '[---]'))
        then
        begin
          Axes := 'VALAXENIV1=[---];';
          Etab := '((QBR_VALEURAXE is null) OR (QBR_VALEURAXE=""))';
        end
        else
        begin
          Axes := 'VALAXENIV1='+Etab+';';
          Etab := 'QBR_VALEURAXE="'+Etab+'"';
        end;

        //Vérification si l'axe sélectionné comprend des dates au dernier niveau
        ReqSQL := '';ComaSQL:='';
        For j := TF.GetValue('QBR_NIVEAU') to NivMax-1 do
        begin
          ReqSQL := ReqSQL + 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUDPERE IN(';
          ComaSQL := ComaSQL + ')';
        end;
        ReqSQL := ReqSQL + 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUDPERE IN ("'
                  +IntToSTr(TF.GetValue('QBR_NUMNOEUD'))+'")' + ComaSQL;

        if not ExisteSQL(ReqSQL) then
        begin
          HShowmessage('1;Initialisation impossible;Le noeud sélectionné ne correspond à aucune date.;W;O;', '', '');
          TF.Cancel;
          MoveCurProgressForm('Calendrier : '+IntToStr((i+1))+'/'+IntToStr(TOB(TV.Selected.Data).Detail.Count));
        end
        else
        begin
          Q:=MOPenSql('SELECT '+PrevHisto+' FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
                    '" AND QBR_NUMNOEUDPERE=0 AND '+Etab,'QUFSLARBRE_TOF (CalendAllHisto_OnClick).',true);
          if not Q.eof then
          begin
            Prevu := Q.Fields[0].AsFloat;
            Histo := Q.Fields[1].AsFloat;
          end;
          Ferme(Q);

          { EVI / ToDo : Protection en cas de valeur nulle }
          CalendValid := aglLanceFiche('Q', 'QUFVBPCALENDJOURM', '', '','SESSION='+codeSession+';'
                                       +Axes+';CALEND='+Calendrier+';MODEFICHE=CALENDTYPE;PREVU='+FloatToStr(Prevu)
                                       +';HISTO='+FloatToStr(Histo)+';VALIDATIONAUTO='+ValidationAuto);

          if ValidationAuto='FALSE' then
          begin
            if CalendValid='OK' then TF.Post else TF.Cancel;
          end
          else
          begin
            TF.Post;
            MoveCurProgressForm('Calendrier : '+IntToStr((i+1))+'/'+IntToStr(TOB(TV.Selected.Data).Detail.Count));
          end;
        end;
        SelectedNode.Collapse(false);
        SelectedNode:=SelectedNode.getNextSibling;
        TF.Next;
      end;

      if ValidationAuto='TRUE' then FiniMoveProgressForm;

      deblocageMenu(TFSaisieList(Ecran));
    end
    else TF.Cancel;
    deblocageMenu(TFSaisieList(Ecran));
  end;
end;

procedure TOF_QUFSLARBRE.CalendAllHisto_OnClick(Sender: TObject);
var i,j : integer;
Axes,CalendValid,Etab,PrevHisto,ValidationAuto : hString;
OKCalend:boolean;
SelectedNode : TTreeNode;
DateDebC,DateFinC:TDateTime;
Prevu,Histo:double;
Q:TQuery;
ReqSQL,ComaSQL : string;
begin
  Axes := '';
  Prevu:=0;
  Histo:=0;
  OKCalend := true;
  SelectedNode := TV.Selected.GetFirstChild;

  { EVI / Test des sous-niveaux }
  if existeSQL( 'SELECT ##TOP 1## QBE_CALENDREP FROM QBPDETCALENDREP WHERE QBE_CALENDREP="'+codeSession+'"' ) then
  begin
    HShowmessage('1;Initialisation impossible;Il existe au moins un calendrier défini pour un des sous-niveaux de votre sélection.;W;O;', '', '');
    OKCalend := false;
  end;

  if OKCalend = true then
  begin
    if TOB(TV.Selected.Data).Detail.Count > 1 then
    begin
      if HShowmessage('1;Initialisation;'+IntToStr(TOB(TV.Selected.Data).Detail.Count)+' calendriers vont '+
                      'être générés.#13#10 Etes-vous sûr ?;Q;YN;N;N;', '', '') = Mrno
      then OKCalend := false;
    end
    else
    begin
      if HShowmessage('1;Initialisation;Un calendrier va être généré.#13#10 Etes-vous sûr ?;Q;YN;N;N;', '', '') = Mrno
      then OKCalend := false;
    end;
  end;

  if OKCalend=true then
  begin
    if HShowmessage('1;Calendriers;Voulez-vous que tous les calendriers soient validés automatiquement ?;Q;YN;N;N;', '', '') = MrYes
    then ValidationAuto:='TRUE' else ValidationAuto:='FALSE';

    dateDebFinSesssion(codeSession,DateDebC,DateFinC);
    { EVI / Remplissage de la table BPCUBETMP }
    InitMoveProgressForm(nil,TraduireMemoire('Génération en cours'),TraduireMemoire('Veuillez patienter'), 1 + NbDeMaille(codesession,BPInitialise,0,0), True, True);
    RemplitTableQBPCubeTmpPgi(codeSession,'','',DateDebC,DateFinC,0,0);

    if not MoveCurProgressForm(TraduireMemoire('Génération...')) then
    begin
      FiniMoveProgressForm;
      exit;
    end;

    FiniMoveProgressForm;

    case ValeurAffiche of
        1 : PrevHisto := 'QBR_OP1,QBR_REF1';
        2 : PrevHisto := 'QBR_QTEC,QBR_QTEREF';
        3 : PrevHisto := 'QBR_OP2,QBR_REF2';
        4 : PrevHisto := 'QBR_OP3,QBR_REF3';
        5 : PrevHisto := 'QBR_OP4,QBR_REF4';
        6 : PrevHisto := 'QBR_OP5,QBR_REF5';
        7 : PrevHisto := 'QBR_OP6,QBR_REF6';
    end;

    blocageMenu(TFSaisieList(Ecran));
    TF.First;
    //Parcours les filles pour relever les Valeur Axe
    if ValidationAuto='TRUE' then
    InitMoveProgressForm(nil,TraduireMemoire('Enregistrement des calendriers'),TraduireMemoire('Veuillez patienter'), TOB(TV.Selected.Data).Detail.Count, False, True);

    for i:=0 to TOB(TV.Selected.Data).Detail.Count-1 do
    begin
      TF.State := dsEdit;
      Etab:=TOB(TV.Selected.Data).Detail[i].GetString('QBG_VALEURAXE');

      if ((Etab = '') OR (Etab = '[---]'))
      then
      begin
        Axes := 'VALAXENIV1=[---];';
        Etab := '((QBR_VALEURAXE is null) OR (QBR_VALEURAXE=""))';
      end
      else
      begin
        Axes := 'VALAXENIV1='+Etab+';';
        Etab := 'QBR_VALEURAXE="'+Etab+'"';
      end;

      //Vérification si l'axe sélectionné comprend des dates au dernier niveau
      ReqSQL := '';ComaSQL:='';
      For j := TF.GetValue('QBR_NIVEAU') to NivMax-1 do
      begin
        ReqSQL := ReqSQL + 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUDPERE IN (';
        ComaSQL := ComaSQL + ')';
      end;
      ReqSQL := ReqSQL + 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUDPERE IN ("'
                +IntToSTr(TF.GetValue('QBR_NUMNOEUD'))+'")' + ComaSQL;

      if not ExisteSQL(ReqSQL) then
      begin
        HShowmessage('1;Initialisation impossible;Le noeud sélectionné ne correspond à aucune date.;W;O;', '', '');
        TF.Cancel;
        MoveCurProgressForm('Calendrier : '+IntToStr((i+1))+'/'+IntToStr(TOB(TV.Selected.Data).Detail.Count));
      end
      else
      begin
        Q:=MOPenSql('SELECT '+PrevHisto+' FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
                    '" AND QBR_NUMNOEUDPERE=0 AND '+Etab,'QUFSLARBRE_TOF (CalendAllHisto_OnClick).',true);

        if not Q.eof then
        begin
          Prevu := Q.Fields[0].AsFloat;
          Histo := Q.Fields[1].AsFloat;
        end;
        Ferme(Q);

        { EVI / ToDo : Protection en cas de valeur nulle }
        CalendValid := aglLanceFiche('Q', 'QUFVBPCALENDJOURM', '', '','SESSION='+codeSession+';'
                                     +Axes+';PREVU='+FloatToStr(Prevu)+';HISTO='+FloatToStr(Histo)+';VALIDATIONAUTO='+ValidationAuto);

        if ValidationAuto='FALSE' then
        begin
          if CalendValid='OK' then TF.Post else TF.Cancel;
        end
        else
        begin
          TF.Post;
          MoveCurProgressForm('Calendrier : '+IntToStr((i+1))+'/'+IntToStr(TOB(TV.Selected.Data).Detail.Count));
        end;
      end;
      SelectedNode.Collapse(false);
      SelectedNode:=SelectedNode.getNextSibling;
      TF.Next;
    end;

    if ValidationAuto='TRUE' then FiniMoveProgressForm;

    deblocageMenu(TFSaisieList(Ecran));
  end;
end;


procedure TOF_QUFSLARBRE.CalendHisto_OnClick(Sender: TObject);
var Level,i : integer;
AxeSQL,Axes,CalendValid,LimitTreeH : hString;
OKCalend:boolean;
SelectedNode : TTreeNode;
DateDebC,DateFinC:TDateTime;
Prevu,Histo:double;
ReqSQL,ComaSQL : string;
begin
  Axes := '';
  AxeSQL := '';
  Prevu:=0;
  Histo:=0;
  OKCalend := true;
  Level := TV.Selected.level;
  SelectedNode := TV.Selected;

  if Level > 0 then
  begin
    for i := 1 to (Level-1) do
    begin
      if ((TF.GetValue('QBR_VALAXENIV'+IntToStr(i)) = '') OR (TF.GetValue('QBR_VALAXENIV'+IntToStr(i)) = '[---]')) then
      begin
        AxeSQL := AxeSQL + 'AND QBE_VALAXENIV'+IntToStr(i)+'="[---]"';
        Axes := Axes + 'VALAXENIV'+IntToStr(i)+'=[---];';
        if LimitTreeH = '' then LimitTreeH := '("")' else LimitTreeH := LimitTreeH + '@@' + '("")'
      end
      else
      begin
        AxeSQL := AxeSQL + 'AND QBE_VALAXENIV'+IntToStr(i)+'="'+TF.GetValue('QBR_VALAXENIV'+IntToStr(i))+'"';
        Axes := Axes + 'VALAXENIV'+IntToStr(i)+'='+TF.GetValue('QBR_VALAXENIV'+IntToStr(i))+';';
        if LimitTreeH = '' then LimitTreeH := '("'+TF.GetValue('QBR_VALAXENIV'+IntToStr(i))+'")'
        else LimitTreeH := LimitTreeH + '@@' + '("'+TF.GetValue('QBR_VALAXENIV'+IntToStr(i))+'")'
      end
    end;

    if ((TF.GetValue('QBR_VALEURAXE') = '') OR (TF.GetValue('QBR_VALEURAXE') = '[---]')) then
    begin
      AxeSQL :=  AxeSQL + 'AND QBE_VALAXENIV'+IntToStr(Level)+'="[---]"';
      Axes :=  Axes + 'VALAXENIV'+IntToStr(Level)+'=[---];';
      if LimitTreeH = '' then LimitTreeH := '("")' else LimitTreeH := LimitTreeH + '@@' + '("")'
    end
    else
    begin
      AxeSQL :=  AxeSQL + 'AND QBE_VALAXENIV'+IntToStr(Level)+'="'+TF.GetValue('QBR_VALEURAXE')+'"';
      Axes :=  Axes + 'VALAXENIV'+IntToStr(Level)+'='+TF.GetValue('QBR_VALEURAXE')+';';
      if LimitTreeH = '' then LimitTreeH := '("'+TF.GetValue('QBR_VALEURAXE')+'")'
      else LimitTreeH := LimitTreeH + '@@' + '("'+TF.GetValue('QBR_VALEURAXE')+'")'
    end
  end;

  //Vérification si l'axe sélectionné comprend des dates au dernier niveau
  For i := TF.GetValue('QBR_NIVEAU') to NivMax-1 do
  begin
    ReqSQL := ReqSQL + 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUDPERE IN(';
    ComaSQL := ComaSQL + ')';
  end;
  ReqSQL := ReqSQL + 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUDPERE IN ("'
            +IntToSTr(TF.GetValue('QBR_NUMNOEUD'))+'")' + ComaSQL;
  if not ExisteSQL(ReqSQL) then
  begin
    HShowmessage('1;Initialisation impossible;Le noeud sélectionné ne correspond à aucune date.;W;O;', '', '');
    exit;
  end;

  { EVI / Test du noeud }
  if TOB(TV.Selected.Data).FieldExists('CALEND') then
  begin
    if Tob(TV.Selected.data).GetValue('CALEND') = 'X' then
    begin
      HShowmessage('1;Initialisation impossible;Un calendrier correspondant à votre sélection existe déjà.;W;O;', '', '');
      OKCalend := false;
    end
    else
    begin { EVI / Test des niveaux supérieurs }
      while TV.Selected.Level <> 0 do
      begin
        TV.Selected := TV.Selected.Parent ;
        if TOB(TV.Selected.Data).FieldExists('CALEND') then
        begin
          if Tob(TV.Selected.data).GetValue('CALEND') = 'X' then
          begin
            HShowmessage('1;Initialisation impossible;Il existe au moins un calendrier défini pour un niveau supérieur.;W;O;', '', '');
            OKCalend := false;
            break;
          end;
        end;
      end;
      if OKCalend = true then
      begin { EVI / Test des sous-niveaux }
        if existeSQL( 'SELECT ##TOP 1## QBE_CALENDREP FROM QBPDETCALENDREP WHERE QBE_CALENDREP="'+codeSession+'"'+AxeSQL  ) then
        begin
          HShowmessage('1;Initialisation impossible;Il existe au moins un calendrier défini pour un des sous-niveaux de votre sélection.;W;O;', '', '');
          OKCalend := false;
        end
      end;
    end;
  end;

  if OKCalend = true then
  begin
    blocageMenu(TFSaisieList(Ecran));
    SelectedNode.Selected:=true;

    dateDebFinSesssion(codeSession,DateDebC,DateFinC);
    { EVI / Remplissage de la table BPCUBETMP }
    InitMoveProgressForm(nil,TraduireMemoire('Génération en cours'),TraduireMemoire('Veuillez patienter'), 1 + NbDeMaille(codesession,BPInitialise,0,0), True, True);

    { EVI / Génération du cube avec restriction au noeud sélectionné par l'utilisateur }
    RemplitTableQBPCubeTmpPgi(codeSession,'',LimitTreeH,DateDebC,DateFinC,0,0);

    if not MoveCurProgressForm(TraduireMemoire('Génération...')) then
    begin
      FiniMoveProgressForm;
      TF.Cancel;
      deblocageMenu(TFSaisieList(Ecran));
      exit;
    end;

    FiniMoveProgressForm;

    case ValeurAffiche of
      1 : begin
            Prevu := TF.GetValue('QBR_OP1');
            Histo := TF.GetValue('QBR_REF1');
          end;
      2 : begin
            Prevu := TF.GetValue('QBR_QTEC');
            Histo := TF.GetValue('QBR_QTEREF');
          end;
      3 : begin
            Prevu := TF.GetValue('QBR_OP2');
            Histo := TF.GetValue('QBR_REF2');
          end;
      4 : begin
            Prevu := TF.GetValue('QBR_OP3');
            Histo := TF.GetValue('QBR_REF3');
          end;
      5 : begin
            Prevu := TF.GetValue('QBR_OP4');
            Histo := TF.GetValue('QBR_REF4');
          end;
      6 : begin
            Prevu := TF.GetValue('QBR_OP5');
            Histo := TF.GetValue('QBR_REF5');
          end;
      7 : begin
            Prevu := TF.GetValue('QBR_OP6');
            Histo := TF.GetValue('QBR_REF6');
          end;
    end;

    TF.State := dsEdit;
    CalendValid := aglLanceFiche('Q', 'QUFVBPCALENDJOURM', '', '','SESSION='+codeSession+';'
                                 +Axes+';PREVU='+FloatToStr(Prevu)+';HISTO='+FloatToStr(Histo));
    if CalendValid='OK' then TF.Post else TF.Cancel;
    deblocageMenu(TFSaisieList(Ecran));
  end;
end;

procedure TOF_QUFSLARBRE.CalendType_OnClick(Sender: TObject);
var DateDebC,DateFinC:TDateTime;
    AxeSQL,Axes,codeSql,niveau,Calendrier,ChpValAff,CalendValid,LimitTreeT:hString;
    OKCalend:boolean;
    Prevu,Histo:double;
    Level,i:integer;
    SelectedNode : TTreeNode;
    ReqSQL,ComaSQL : string;
begin
  Axes := '';
  AxeSQL := '';
  Prevu:=0;
  Histo:=0;
  OKCalend := true;
  Level := TV.Selected.level;
  SelectedNode := TV.Selected;

  if Level > 0 then
  begin
    for i := 1 to (Level-1) do
    begin
      if ((TF.GetValue('QBR_VALAXENIV'+IntToStr(i)) = '') OR (TF.GetValue('QBR_VALAXENIV'+IntToStr(i)) = '[---]')) then
      begin
        AxeSQL := AxeSQL + 'AND QBE_VALAXENIV'+IntToStr(i)+'="[---]"';
        Axes := Axes + 'VALAXENIV'+IntToStr(i)+'=[---];';
        if LimitTreeT = '' then LimitTreeT := '("")' else LimitTreeT := LimitTreeT + '@@' + '("")'
      end
      else
      begin
        AxeSQL := AxeSQL + 'AND QBE_VALAXENIV'+IntToStr(i)+'="'+TF.GetValue('QBR_VALAXENIV'+IntToStr(i))+'"';
        Axes := Axes + 'VALAXENIV'+IntToStr(i)+'='+TF.GetValue('QBR_VALAXENIV'+IntToStr(i))+';';
        if LimitTreeT = '' then LimitTreeT := '("'+TF.GetValue('QBR_VALAXENIV'+IntToStr(i))+'")'
        else LimitTreeT := LimitTreeT + '@@' + '("'+TF.GetValue('QBR_VALAXENIV'+IntToStr(i))+'")'
      end
    end;
    if Tob(TV.Selected.data).GetValue('SOUSNIVEAU')<> 0 then
    begin
      if ((TF.GetValue('QBR_VALEURAXE') = '') OR (TF.GetValue('QBR_VALEURAXE') = '[---]')) then
      begin
        AxeSQL :=  AxeSQL + 'AND QBE_VALAXENIV'+IntToStr(Level)+'="[---]"';
        Axes :=  Axes + 'VALAXENIV'+IntToStr(Level)+'=[---];';
        if LimitTreeT = '' then LimitTreeT := '("")' else LimitTreeT := LimitTreeT + '@@' + '("")'
      end
      else
      begin
        AxeSQL :=  AxeSQL + 'AND QBE_VALAXENIV'+IntToStr(Level)+'="'+TF.GetValue('QBR_VALEURAXE')+'"';
        Axes :=  Axes + 'VALAXENIV'+IntToStr(Level)+'='+TF.GetValue('QBR_VALEURAXE')+';';
        if LimitTreeT = '' then LimitTreeT := '("'+TF.GetValue('QBR_VALEURAXE')+'")'
        else LimitTreeT := LimitTreeT + '@@' + '("'+TF.GetValue('QBR_VALEURAXE')+'")'
      end
    end;
  end;

  //Vérification si l'axe sélectionné comprend des dates au dernier niveau
  For i := TF.GetValue('QBR_NIVEAU') to NivMax-1 do
  begin
    ReqSQL := ReqSQL + 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUDPERE IN(';
    ComaSQL := ComaSQL + ')';
  end;
  ReqSQL := ReqSQL + 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUDPERE IN ("'
            +IntToSTr(TF.GetValue('QBR_NUMNOEUD'))+'")' + ComaSQL;
  if not ExisteSQL(ReqSQL) then
  begin
    HShowmessage('1;Initialisation impossible;Le noeud sélectionné ne correspond à aucune date.;W;O;', '', '');
    exit;
  end;

  { EVI / Test du noeud }
  if TOB(TV.Selected.Data).FieldExists('CALEND') then
  begin
    if Tob(TV.Selected.data).GetValue('CALEND') = 'X' then
    begin
      HShowmessage('1;Initialisation impossible;Un calendrier correspondant à votre sélection existe déjà.;W;O;', '', '');
      OKCalend := false;
    end
    else
    begin { EVI / Test des niveaux supérieurs }
      while TV.Selected.Level <> 0 do
      begin
        TV.Selected := TV.Selected.Parent ;
        if TOB(TV.Selected.Data).FieldExists('CALEND') then
        begin
          if Tob(TV.Selected.data).GetValue('CALEND') = 'X' then
          begin
            HShowmessage('1;Initialisation impossible;Il existe au moins un calendrier défini pour un niveau supérieur.;W;O;', '', '');
            OKCalend := false;
            break;
          end;
        end;
      end;
      if OKCalend = true then
      begin { EVI / Test des sous-niveaux }
        if existeSQL( 'SELECT ##TOP 1## QBE_CALENDREP FROM QBPDETCALENDREP WHERE QBE_CALENDREP="'+codeSession+'"'+AxeSQL  ) then
        begin
          HShowmessage('1;Initialisation impossible;Il existe au moins un calendrier défini pour un des sous-niveaux de votre sélection.;W;O;', '', '');
          OKCalend := false;
        end
      end;
    end;
  end;

  if OKCalend = true then
  begin
    blocageMenu(TFSaisieList(Ecran));
    SelectedNode.Selected:=true;

    Calendrier:=AglLanceFiche('Q','QUFMBPCALENDREP','','','CALENDTYPE');
    if (Calendrier <> 'CANCEL') AND (Calendrier <> '') then
    begin

      dateDebFinSesssion(codeSession,DateDebC,DateFinC);
      TF.State := dsEdit;
      { EVI / Remplissage de la table BPCUBETMP }
      InitMoveProgressForm(nil,TraduireMemoire('Génération en cours'),TraduireMemoire('Veuillez patienter'), 1 + NbDeMaille(codesession,BPInitialise,0,0), True, True);
      { EVI / Génération du cube avec restriction au noeud sélectionné par l'utilisateur }
      RemplitTableQBPCubeTmpPgi(codeSession,'',LimitTreeT,DateDebC,DateFinC,0,0);

      if not MoveCurProgressForm(TraduireMemoire('Génération...')) then
      begin
        FiniMoveProgressForm;
        TF.Cancel;
        exit;
      end;

      FiniMoveProgressForm;

      ChpValAff:='QBR_OP1';
      case ValeurAffiche of
        1 : ChpValAff:='QBR_OP1';
        2 : ChpValAff:='QBR_QTEC';
        3 : ChpValAff:='QBR_OP2';
        4 : ChpValAff:='QBR_OP3';
        5 : ChpValAff:='QBR_OP4';
        6 : ChpValAff:='QBR_OP5';
        7 : ChpValAff:='QBR_OP6';
      end;

      case ValeurAffiche of
        1 : begin
              Prevu := TF.GetValue('QBR_OP1');
              Histo := TF.GetValue('QBR_REF1');
            end;
        2 : begin
              Prevu := TF.GetValue('QBR_QTEC');
              Histo := TF.GetValue('QBR_QTEREF');
            end;
        3 : begin
              Prevu := TF.GetValue('QBR_OP2');
              Histo := TF.GetValue('QBR_REF2');
            end;
        4 : begin
              Prevu := TF.GetValue('QBR_OP3');
              Histo := TF.GetValue('QBR_REF3');
            end;
        5 : begin
              Prevu := TF.GetValue('QBR_OP4');
              Histo := TF.GetValue('QBR_REF4');
            end;
        6 : begin
              Prevu := TF.GetValue('QBR_OP5');
              Histo := TF.GetValue('QBR_REF5');
            end;
        7 : begin
              Prevu := TF.GetValue('QBR_OP6');
              Histo := TF.GetValue('QBR_REF6');
            end;
      end;

      niveau:=TF.GetValue('QBR_NIVEAU');
      codeSql:=' SELECT '+ChpValAff+' FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NIVEAU="'+niveau+'" ';
      for i:=1 to VALEURI(niveau)-1 do
      codeSql:=codeSql+' AND QBR_VALAXENIV'+intToStr(i)+'="'+TF.GetValue('QBR_VALAXENIV'+intToStr(i))+'" ';
      codeSql:=codeSql+' AND QBR_VALEURAXE="';

      CalendValid := aglLanceFiche('Q', 'QUFVBPCALENDJOURM', '', '','SESSION='+codeSession+';'
                     +Axes+';CALEND='+Calendrier+';MODEFICHE=CALENDTYPE;PREVU='+FloatToStr(Prevu)+';HISTO='+FloatToStr(Histo));
      if CalendValid='OK' then TF.Post else TF.Cancel;

    end
    else TF.Cancel;
    deblocageMenu(TFSaisieList(Ecran));
  end;
end;

procedure TOF_QUFSLARBRE.Refresh_OnClick(Sender: TObject);
begin
  SetImages(TV,nil);
  TV.Refresh;
end;

procedure TOF_QUFSLARBRE.Expand_OnClick(Sender: TObject);
begin
  if TV.Selected.Expanded = true then TV.Selected.Collapse(True)
  else TV.Selected.Expand(True);
end;

procedure TOF_QUFSLARBRE.ExpandAll_OnClick(Sender: TObject);
begin
  TV.FullExpand;
end;

procedure TOF_QUFSLARBRE.CollapseAll_OnClick(Sender: TObject);
begin
  TV.FullCollapse;
end;

procedure TOF_QUFSLARBRE.SetImages (TV: TTreeView; TN: TTreeNode) ;
var tn1: TTreeNode;
begin
  if (TV = nil) and (TN = nil) then Exit;
  if TV <> nil then tn1 := TV.TopItem
  else if TN.Count <> 0 then tn1 := TN.getFirstChild
  else tn1 := TN;

  while tn1 <> nil do
  begin
    if tn1.text <> 'DUMMY' then
    begin
      if TOB(tn1.Data).FieldExists('CALEND') then
      begin
        if TOB(tn1.Data).GetValue('CALEND') = 'X' then
        begin
          if TOB(tn1.Data).GetValue('QBG_VALBLOQUE') = 'X' then  tn1.ImageIndex := 3
          else tn1.ImageIndex := 1;
        end
        else
        begin
          if TOB(tn1.Data).GetValue('QBG_VALBLOQUE') = 'X' then  tn1.ImageIndex := 2
          else tn1.ImageIndex := 0;
        end;
      end
      else tn1.ImageIndex := 4;
      if tn1.count <> 0 then SetImages (nil, tn1) ;
    end;
    if (TV <> nil) or (TN.Count <> 0) then tn1 := tn1.GetNextSibling
    else tn1 := nil;
  end;
end;

procedure TOF_QUFSLARBRE.TV_OnExpanded(Sender: TObject; TN: TTreeNode);
begin
  SetImages(TV,TN);
end ;

procedure TOF_QUFSLARBRE.TV_OnGetSelectedIndex(Sender: TObject; TN: TTreeNode);
begin
  if TOB(TN.Data).FieldExists('CALEND') then
  begin
    if TOB(TN.Data).GetValue('CALEND') = 'X' then
    begin
      if TOB(TN.Data).GetValue('QBG_VALBLOQUE') = 'X' then TN.SelectedIndex := 3
      else TN.SelectedIndex := 1;
    end
    else
    begin
      if TOB(TN.Data).GetValue('QBG_VALBLOQUE') = 'X' then TN.SelectedIndex := 2
      else TN.SelectedIndex := 0;
    end;
  end
  else TN.SelectedIndex := 4;
  if TOB(TV.Selected.Data) <> nil then
  begin
    if TOB(TV.Selected.Data).FieldExists('CALEND') then
    begin
      if TOB(TV.Selected.Data).GetValue('CALEND') = 'X'
      then TToolBarButton97(GetControl( 'BTNCALREP' )).visible := true
      else TToolBarButton97(GetControl( 'BTNCALREP' )).visible := false;
    end
    else TToolBarButton97(GetControl( 'BTNCALREP' )).visible := false;
  end
  else TToolBarButton97(GetControl( 'BTNCALREP' )).visible := false;
end ;

procedure TOF_QUFSLARBRE.AfterBParamClick(Sender: TObject);
begin
  TFSaisieList(Ecran).BParamListe.Click;
  PageChange(Self);
  if ContextBP=3 then DoSetNavigate(Self);
end;

procedure TOF_QUFSLARBRE.CreateSousNiveauAuto_OnClick(Sender: TObject);
var retour:hString;
SelectedNode:TTreeNode;
OKTraitement:boolean;
DateDebC,DateFinC:TDateTime;
begin
  OKTraitement:=true;
  dateDebFinSesssion(codeSession,DateDebC,DateFinC);
  numnoeudT:=TF.GetValue('QBR_NUMNOEUD');
  numnoeudpereT:=TF.GetValue('QBR_NUMNOEUDPERE');
  niveauT:=TF.GetValue('QBR_NIVEAU');
  SelectedNode := TV.Selected;

  if TOB(TV.Selected.Data).FieldExists('QBG_VALBLOQUE') then
  begin
    if Tob(TV.Selected.data).GetValue('QBG_VALBLOQUE') = 'X' then
    begin
      HShowmessage('1;Traitement impossible;Le noeud courant est bloqué.;W;O;', '', '');
      OKTraitement := false;
    end
    else
    begin
      while TV.Selected.Level <> 0 do
      begin
        TV.Selected := TV.Selected.Parent ;
        if TOB(TV.Selected.Data).FieldExists('QBG_VALBLOQUE') then
        begin
          if Tob(TV.Selected.data).GetValue('QBG_VALBLOQUE') = 'X' then
          begin
            HShowmessage('1;Traitement impossible;Il existe au moins un noeud de niveau supérieur bloqué.;W;O;', '', '');
            OKTraitement := false;
            break;
          end;
        end;
      end;
    end;
  end;
  SelectedNode.Selected:=true;

  if OKTraitement then
  begin
    retour:=aglLanceFiche('Q', 'QUFVBPSUBNIVAUTO', '', '',
                          'SESSION='+codeSession+';NUMNOEUD='+numnoeudT+';NIVEAU='+niveauT+
                          ';DATEDEB='+DateTimeToStr(DateDebC)+';DATEFIN='+DateTimeToStr(DateFinC)+
                          ';CODEAXE='+TF.GetValue('QBR_CODEAXE')+';VALEURCODEAXE='+TF.GetValue('QBR_VALEURAXE')+
                          ';CODEAXEPREC='+TabCodeAxe[VALEURI(niveauT)]+';FONCTION=CREATESOUSNIVO');
    if retour<>'' then
    begin
      CreateSousNiveauAuto(retour,codeSession,niveauT,numnoeudT);
      TF.RefreshLignes;
      TF.State := dsEdit;
      TF.Post;
    end;
  end;
end;

procedure TOF_QUFSLARBRE.ModifSousNiveauAuto_OnClick(Sender: TObject);
var retour:hString;
SelectedNode:TTreeNode;
OKTraitement:boolean;
DateDebC,DateFinC:TDateTime;
begin
  //Protection si entête de l'arbre
  if Tob(TV.Selected.data).GetValue('QBG_NIVEAU')=0 then exit;
  OKTraitement:=true;
  dateDebFinSesssion(codeSession,DateDebC,DateFinC);
  numnoeudT:=TF.GetValue('QBR_NUMNOEUD');
  numnoeudpereT:=TF.GetValue('QBR_NUMNOEUDPERE');
  niveauT:=TF.GetValue('QBR_NIVEAU');
  SelectedNode := TV.Selected;

  if TOB(TV.Selected.Data).FieldExists('QBG_VALBLOQUE') then
  begin
    if Tob(TV.Selected.data).GetValue('QBG_VALBLOQUE') = 'X' then
    begin
      HShowmessage('1;Traitement impossible;Le noeud courant est bloqué.;W;O;', '', '');
      OKTraitement := false;
    end
    else
    begin
      while TV.Selected.Level <> 0 do
      begin
        TV.Selected := TV.Selected.Parent ;
        if TOB(TV.Selected.Data).FieldExists('QBG_VALBLOQUE') then
        begin
          if Tob(TV.Selected.data).GetValue('QBG_VALBLOQUE') = 'X' then
          begin
            HShowmessage('1;Traitement impossible;Il existe au moins un noeud de niveau supérieur bloqué.;W;O;', '', '');
            OKTraitement := false;
            break;
          end;
        end;
      end;
    end;
  end;
  SelectedNode.Selected:=true;

  if OKTraitement then
  begin
    retour:=aglLanceFiche('Q', 'QUFVBPSUBNIVAUTO', '', '',
                          'SESSION='+codeSession+';NUMNOEUD='+numnoeudT+';NIVEAU='+niveauT+
                          ';DATEDEB='+DateTimeToStr(DateDebC)+';DATEFIN='+DateTimeToStr(DateFinC)+
                          ';CODEAXE='+TF.GetValue('QBR_CODEAXE')+';VALEURCODEAXE='+TF.GetValue('QBR_VALEURAXE')+
                          ';CODEAXEPREC='+TabCodeAxe[VALEURI(niveauT)]+';FONCTION=MODIFSOUSNIVO');

    if retour<>'' then
    begin
      ModifSousNiveauAuto(retour,codeSession,codeaxeT,TF.GetValue('QBR_VALEURAXE'),niveauT,numnoeudT);
      TF.RefreshLignes;
      TF.State := dsEdit;
      TF.Post;
    end;
  end;
end;

procedure TOF_QUFSLARBRE.InsertNewSalary_OnClick(Sender: TObject);
var i:integer;
retour:hString;
begin
  retour:=aglLanceFiche('Q', 'QUFVBPNEWVAL', '', '',
                        'SESSION='+codeSession+';NUMNOEUD='+numnoeudT+';NIVEAU='+niveauT+
                        ';CODEAXE='+TF.GetValue('QBR_CODEAXE')+';VALEURCODEAXE='+TF.GetValue('QBR_VALEURAXE')+
                        ';CODEAXEPREC='+TabCodeAxe[VALEURI(niveauT)]+';FONCTION=MODIFSOUSNIVO');

  //ToDo: rafraichir l'arbre si on rejoute un noeud de niveau 1
  if retour<>'' then
  begin
    InsertNewVal(retour,codeSession);
    TV.Select(TV.Items[1],[]);
    TF.RefreshLignes;
    TF.First;
    OldCodeAxe:='0';
    for i:=1 to TF.LaGrid.RowCount-1 do
    begin
      TF.State := dsEdit;
      TF.Post;
      TV.Selected.Collapse(false);
      TF.Next;
    end;
  end;
end;

procedure TOF_QUFSLARBRE.ModifNodes_OnClick(Sender: TObject);
var i:integer;
retour:hString;
begin
  retour:=aglLanceFiche('Q', 'QUFVBPEVOLUTION2', '', '',
                        'SESSION='+codeSession+';NUMNOEUD='+numnoeudT+';NIVEAU='+niveauT+
                        ';CODEAXE='+TF.GetValue('QBR_CODEAXE')+';VALEURCODEAXE='+TF.GetValue('QBR_VALEURAXE')+
                        ';CODEAXEPREC='+TabCodeAxe[VALEURI(niveauT)]+';FONCTION=MODIFSOUSNIVO');

  if retour<>'' then
  begin
    EvolutionTwo(retour,codeSession);
    TV.Select(TV.Items[1],[]);
    TF.RefreshLignes;
    TF.First;
    OldCodeAxe:='0';
    for i:=1 to TF.LaGrid.RowCount-1 do
    begin
      TF.State := dsEdit;
      TF.Post;
      TV.Selected.Collapse(false);
      TF.Next;
    end;
  end;
end;

procedure TOF_QUFSLARBRE.DoAfterLoadGrid(Sender: TObject);
var j: integer;
    Cn: THNumedit;
Begin
  if TF.PanelCumul<>nil then
  begin
    for j := 0 to TF.PanelCumul.ControlCount - 1 do
    begin
      if TF.PanelCumul.Controls[j] is THNumEdit then
      begin
        Cn := THNumEdit(TF.PanelCumul.Controls[j]); //Cn.Tag = numéro de la colonne
        Cn.OnChange := DoOnAffecteCumul;
        Cn.NumericType := ntDecimal 
      end;
    end;
  end;
end ;

procedure TOF_QUFSLARBRE.DoOnAffecteCumul(Sender: TObject);
var  Cn: THNumEdit;
  OldChange: TNotifyevent;
  ETPValues:boolean;
  i:integer;
begin
  ETPValues:=false;
  if TF.GetValue('QBR_CODEAXE')='DELAI' then exit;
  Cn := (Sender as THNumEdit);
  OldChange := Cn.OnChange;
  for i:=1 to 14 do
  begin
    if TF.LaGrid.ColNames[Cn.Tag]=TabRefSQLValues[i] then
    begin
      ETPValues:=true;
      break;
    end
  end;
  if ETPValues=false then exit;
  Cn.OnChange := nil; //Mettre à nil avant d'affecter pour éviter la réentrance
  if TF.TOBFiltre.detail.count>0 then Cn.Value := Cn.Value / NbIntervalle;
  Cn.OnChange := OldChange;
end;



Initialization
  registerclasses ( [ TOF_QUFSLARBRE ] ) ;
end.

