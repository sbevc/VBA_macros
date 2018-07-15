Attribute VB_Name = "mdCustomizeMenu"
Sub CustomizeRightClickMenu()

    Dim ContextMenu 'Menu click en celdas
    Dim SubMenu1    'Menu MyMacros
    Dim SubMenu1_1  'Menu MyMacros/General
    Dim SubMenu1_2  'Menu MyMacros/GAMMA
    
    ' Delete the controls first to avoid duplicates.
    Call DeleteFromCellMenu
    
    '-----MENU GENERAL-----'
    Set ContextMenu = Application.CommandBars("Cell")
    
    
    '-----SUBMENU MYMACROS-----'
    Set SubMenu1 = ContextMenu.Controls.Add(Type:=msoControlPopup, before:=1)
    
    With SubMenu1
        .Caption = "MyMacros"
        .Tag = "My_Cell_Control_Tag"
    End With
    
    
    '-----MYMACROS/GENERAL"-----'
    Set SubMenu1_1 = SubMenu1.Controls.Add(Type:=msoControlPopup, before:=1)
    With SubMenu1_1
        .Caption = "General"
         With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "copyUnique"
            .FaceId = 2112
            .Caption = "Copiar valores únicos"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "FixNums"
            .FaceId = 2112
            .Caption = "Fix Nums"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "CopyCondRecNo"
            .FaceId = 2112
            .Caption = "Copiar CondRecNo"
        End With
    End With
        
    Set SubMenu1_2 = SubMenu1.Controls.Add(Type:=msoControlPopup, before:=2)
    With SubMenu1_2
        .Caption = "VIS"
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "show_ufrm"
            .FaceId = 1763
            .Caption = "To VIS"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "getlisting"
            .FaceId = 1763
            .Caption = "Listing"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "getVariantByVAN"
            .FaceId = 2112
            .Caption = "Variante con VAN"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "getVariantByGrouping"
            .FaceId = 2112
            .Caption = "Variante con Grouping"
        End With
    End With
    
    
    '-----MYMACROS/GAMMA"-----'
    Set SubMenu1_3 = SubMenu1.Controls.Add(Type:=msoControlPopup, before:=3)
    
    With SubMenu1_3
        .Caption = "GAMMA"
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "cmdDesproteger_Click"
            .FaceId = 902
            .Caption = "Format MACROGLO"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "FixGamma"
            .FaceId = 902
            .Caption = "FixGamma"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "getGammaSites"
            .FaceId = 902
            .Caption = "Gamma Sites"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "buildDesafich"
            .FaceId = 902
            .Caption = "Desfich art|site"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "MACROGLO_title"
            .FaceId = 902
            .Caption = "MACROGLO title"
        End With
    End With

    
End Sub

Sub DeleteFromCellMenu()
    Dim ContextMenu As CommandBar
    Dim ctrl As CommandBarControl

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Delete the custom controls with the Tag : My_Cell_Control_Tag.
    For Each ctrl In ContextMenu.Controls
        If ctrl.Tag = "My_Cell_Control_Tag" Then
            ctrl.Delete
        End If
    Next ctrl

    ' Delete the custom built-in Save button.
    On Error Resume Next
    ContextMenu.FindControl(ID:=3).Delete
    On Error GoTo 0
End Sub
