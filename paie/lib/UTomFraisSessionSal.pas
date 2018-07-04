{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 17/07/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : FRAISSESSIONSAL (FRAISSESSIONSAL)
Mots clefs ... : TOM;FRAISSESSIONSAL
*****************************************************************}
Unit UTomFraisSessionSal;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
       eFiche,eFichList,     
{$ENDIF}
     sysutils,HCtrls,UTOM,UTob,uTableFiltre,
     SaisieList;

Type
  TOM_FRAISSESSIONSAL = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
    TF: TTableFiltre;
    procedure TVClick(Sender:TObject);
    end ;

Implementation

procedure TOM_FRAISSESSIONSAL.TVClick(Sender:TObject);
var Stage,Session,Millesime:String;
    Q:TQuery;
begin
Stage:=TF.TOBFiltre.GetValue('PFO_CODESTAGE');
Session:=IntToStr(TF.TOBFiltre.GetValue('PFO_ORDRE'));
Millesime:=TF.TOBFiltre.GetValue('PFO_MILLESIME');
Q:=OpenSQL('SELECT PSS_LIBELLE,PST_LIBELLE,PST_LIBELLE1,PSS_DATEDEBUT,PSS_DATEFIN,PST_DUREESTAGE,PST_JOURSTAGE'+
           ',PSS_NBRESTAGPREV FROM SESSIONSTAGE'+
           ' LEFT JOIN STAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME'+
           ' WHERE PSS_CODESTAGE="'+Stage+'" AND PSS_MILLESIME="'+Millesime+'" AND PSS_ORDRE="'+Session+'"',True);
if not Q.eof then
   begin
   SetControlText('STAGE',Q.FindField('PST_LIBELLE').AsString);
   SetControlText('SESSION',Q.FindField('PSS_LIBELLE').AsString);
   SetControlText('DATEDEBUT',Q.FindField('PSS_DATEDEBUT').AsString);
   SetControlText('DATEFIN',Q.FindField('PSS_DATEFIN').AsString);
   SetControlText('NBRESTAGPREV',IntToStr(Q.FindField('PSS_NBRESTAGPREV').AsInteger));     //DB2
   end;
Ferme(Q);
end;

procedure TOM_FRAISSESSIONSAL.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_FRAISSESSIONSAL.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_FRAISSESSIONSAL.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_FRAISSESSIONSAL.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_FRAISSESSIONSAL.OnLoadRecord ;
begin
  Inherited ;

end ;

procedure TOM_FRAISSESSIONSAL.OnChangeField ( F: TField ) ;
begin
  Inherited ;
TF.LaGrid.Cols[1].Text:='Test';
end ;

procedure TOM_FRAISSESSIONSAL.OnArgument ( S: String ) ;
begin
  Inherited ;
  if (Ecran<>nil) and (Ecran is TFSaisieList ) then
    TF := TFSaisieList(Ecran).LeFiltre;
 THTreeView(GetControl( 'TreeEntete' )).OnClick := TVClick;
SetControlVisible('PanPied',True);
SetControlVisible('PCPied',True);
SetControlVisible('Page1',False);
SetControlProperty('Page3','TabVisible',False);
SetControlVisible('Page2',False);
SetControlProperty('Page2','TabVisible',False);
SetControlVisible('Page3',True);
SetControlProperty('Page3','TabVisible',true);
end ;

procedure TOM_FRAISSESSIONSAL.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_FRAISSESSIONSAL.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_FRAISSESSIONSAL ] ) ; 
end.
