@{
    # Leave DefaultWorkDir empty on purpose. Run-Snap.ps1 will remember the last WorkDir in snap_session.json.
    DefaultWorkDir = ''
    DefaultOwner   = '厳'

    Window = @{
        Width    = 1050
        Height   = 761
        CropPx   = 6
        NoResize = $false
    }

    Timing = @{
        ActionWaitMs = 500
        ResultWaitSec = 2
        ResultWaitMs  = 500
    }

    Hm = @{
        TabsToCorrelid     = 1
        TabsBackFromSearch = 1
        TabsBackToInput    = 4
    }

    Mq = @{
        TabsToInquiry  = 1
        TabsToCorrelid = 4
    }

    Scripts = @{
        GenerateMapping = 'Generate-HostOpenMapping.ps1'
        Excel           = 'ExcelSnap.ps1'
        Hm              = 'HmSnap.ps1'
        Mq              = 'MqSnap.ps1'
        Jenkins         = 'JenkinsSnap.ps1'
        Crop            = 'Crop-Snap.ps1'
        Common          = 'Common.ps1'
    }

    PhaseOrder = @(
        @{ Key='Mapping';        Field='';                    Label='mapping 生成 / 更新' }
        @{ Key='Excel';          Field='Excel_snap';           Label='Excel 証跡' }
        @{ Key='HmGift';         Field='GIFT_HM_snap';         Label='GIFT HM 証跡' }
        @{ Key='MqGift';         Field='GIFT_MQ_snap';         Label='GIFT MQ 証跡' }
        @{ Key='JenkinsGift';    Field='GIFT_Jenkins_snap';    Label='GIFT Jenkins 証跡 + DL' }
        @{ Key='NoGfix';         Field='GIFT_noGfixfile_snap'; Label='GIFT no-GFIX 証跡' }
        @{ Key='HmGfix';         Field='GFIX_HM_snap';         Label='GFIX HM 証跡' }
        @{ Key='JenkinsGfix';    Field='GFIX_Jenkins_snap';    Label='GFIX Jenkins 証跡 + DL' }
        @{ Key='Df';             Field='DF_snap';              Label='DF 証跡（未実装 placeholder）' }
    )
}
