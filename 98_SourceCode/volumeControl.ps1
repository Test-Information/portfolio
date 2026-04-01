# ==============================================================================================================
# プログラム名：【音量調節 スクリプト】
# 【説明】
# 自宅環境で使用しているスクリプトです。
# 品質に問題がないことを確認したつもりではありますが、念のため実行前にコードチェックのうえ実行していただけますと幸いです。
#
# 【使い方】
# 1)引数なしで実行するたびに音量を大、小を切り返る。
# 2)第一引数あり実行した場合、引数の音量に変更する。
# 【作成情報】
# 作成者  ：神部 慶太
# 作成日  ：2026/3/29
# 
# 【前提条件】
# 以下の外部モジュールをインストールしておく。
# Install-Module -Name AudioDeviceCmdlets -Scope CurrentUser
# ==============================================================================================================

# 現在の音量を取得
$current = Get-AudioDevice -PlaybackVolume
$vol = [int]$current.TrimEnd('%')

try
{
    [int]$volAft = $args[0] #変更後音量
    $dbg =$args[1] #デバッグモード指定
    Write-Host "---------------------"
    Write-Host "変更先音量：$($volAft)"
    Write-Host "デバッグモード：$($dbg)"
    Write-Host "引数受取処理完了"
}
catch
{
    Write-Host "---------------------"
    Write-Host "$($args[0].GetType().Name)：$($args[0])"
    Write-Host "$($args[1].GetType().Name)：$($args[1])"
    Write-Host "引数受取処理に失敗しました。"
    pause
    exit
}

Write-Host "---------------------"
#指定音量判定
if($volAft -eq $null)
{
    # 現在の音量を判定
    if ($vol -gt 50) 
    {
        #音量変更(小)
        Set-AudioDevice -PlaybackVolume 30
        Write-Host "音量をデフォルト値に変更しました。"
    }
    else 
    {
        #音量変更(大)
        Set-AudioDevice -PlaybackVolume 60
        Write-Host "音量をデフォルト値に変更しました。"
    }
}
else
{
    #指定音量の判定
    if ($volAft -ge 0 -and $volAft -le 70)
    {
        Set-AudioDevice -PlaybackVolume $args[0]
        Write-Host "指定された音量に変更しました。$($volAft)"
    }
    elseif ($volAft -ge 71 -and $volAft -le 100)
    {
        $upperSnd = 70
        Set-AudioDevice -PlaybackVolume $upperSnd
        Write-Host "音量が大きいため$($upperSnd)に調整します。"
    }
    else
    {
        Write-Host "指定された音量に変更できません。"
        Write-Host "$($volAft.GetType().Name)：$($volAft)"
    }
}

Write-Host "---------------------"
#デバッグモード判定
if ($dbg -eq "d" -Or  $dbg -eq "D")
{
    pause
}

