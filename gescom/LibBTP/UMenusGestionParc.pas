unit UMenusGestionParc;

interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     Hent1,
		 Menus,
		 HCtrls,
     HMsgBox,
     SysUtils,
     AglInit,
     forms,
     UtilPGI,
     EntGC,
     uTob;


function GestionMenuParcMateriel(Numtag : integer) : integer;

implementation

Uses BtPlanning;

function GestionMenuParcMateriel(Numtag : integer) : integer;
Var OldValue  : boolean;
begin

	if not VH_GC.BTSeriaParcMateriel then
  begin
  	PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Gestion Parc Matériel');
    result := -1;
    exit;
  end;

  result := 0;

  Case Numtag of
    // 
    329110 : AGLLanceFiche('BTP','BTFAMILLEMAT_MUL','','','');
    329120 : AGLLanceFiche('BTP','BTMATERIEL_MUL','','','');
    329130 : AGLLanceFiche('BTP','BTTYPEACTION_MUL','','','TYPEACTION=PMA');
    329140 : Agllancefiche('BTP','BTPARAMPLAN_MUL', 'HPP_MODEPLANNING=PMA', '', 'TYPEPLANNING=PMA') ;
    //
    329210 : AGLLanceFiche('BTP','BTAFFECTATION_MUL','','','');
    329220 : AGLLanceFiche('BTP','BTEVENTMAT_MUL','','','');
    //
    329230 :
    begin
      OldValue := V_Pgi.ZoomOle ;
      V_Pgi.ZoomOle := True ;
      try
        Saisieplanning ('PMA', 0, taModif) ;
      finally
        V_Pgi.ZoomOle := OldValue ;
      end;
    end;
    329250 :
    begin
      OldValue := V_Pgi.ZoomOle ;
      V_Pgi.ZoomOle := True ;
      try
        Saisieplanning ('PMA', 0, taConsult) ;
      finally
        V_Pgi.ZoomOle := OldValue ;
      end;
    end;

    329240 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=BCM;GP_VENTEACHAT=ACH','','MODIFICATION') ;
    //
    329310 : AGLLanceFiche('BTP','BTLISTMATFAM','','','') ;
    329320 : AGLLanceFiche('BTP','BTFICHEMATCONFIE','','','') ;


  else HShowMessage('2;?caption?;'+'Fonction non disponible : '+';W;O;O;O;','Gestion Parc Matériels',IntToStr(Numtag)) ;
  end;


end;


end.
