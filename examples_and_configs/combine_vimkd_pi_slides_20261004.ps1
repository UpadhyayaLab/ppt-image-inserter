# combine_vimkd_pi_slides_20261004.ps1
# One deck for the PI (FF 2026-10-04): three single-slide decks, in this order:
#   1. Vim_centrosome_profile_20261004.pptx          ("Vim profiles" session)
#   2. VimKD_centrosome_polarization_20261004.pptx   (insert_vimkd_centrosome_slide.py)
#   3. VimKD_live_siCtrl_cell14_XZ_MIP_20261004.pptx (insert_vimkd_live_cell14_movie_slide.py)
# Sources are opened read-only / inserted from file, never modified. Movie keeps playing
# (InsertFromFile copies the embedded media). Renders go to renders\VimKD_PI_summary_slideNN.png.
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File combine_vimkd_pi_slides_20261004.ps1

$dir  = 'K:\FF\PPT\PPT_autogeneration\Fixed Jurkats, Miscellaneous\VimentinKD'
$srcs = @('Vim_centrosome_profile_20261004.pptx',                 # 1
          'VimKD_centrosome_polarization_20261004.pptx',          # 2
          'VimKD_pericentrin_pooled_and_apr6_20261004.pptx',      # 3-6 (Apr 6, May 20, Jun 28, pooled), FF / AU 2026-10-04
          'VimKD_live_siCtrl_cell14_XZ_MIP_20261004.pptx',        # 7
          'VimKD_live_centrosome_synapse_distance_20261004.pptx') # 8
$out  = Join-Path $dir 'VimKD_PI_summary_20261004.pptx'

$app = New-Object -ComObject PowerPoint.Application
$first = $app.Presentations.Open((Join-Path $dir $srcs[0]), -1, 0, 0)   # ReadOnly, no window
$first.SaveCopyAs($out)
$first.Close()

$deck = $app.Presentations.Open($out, 0, 0, 0)
for ($i = 1; $i -lt $srcs.Count; $i++) {
    $src = $app.Presentations.Open((Join-Path $dir $srcs[$i]), -1, 0, 0)  # read-only, to count its slides
    $ns = $src.Slides.Count; $src.Close()
    [void]$deck.Slides.InsertFromFile((Join-Path $dir $srcs[$i]), $deck.Slides.Count, 1, $ns)
}
$deck.Save()
$rdir = Join-Path $dir 'renders'
if (-not (Test-Path $rdir)) { New-Item -ItemType Directory -Path $rdir | Out-Null }
Get-ChildItem $rdir -Filter 'VimKD_PI_summary_slide*.png' | Remove-Item   # this script's own renders; avoids stale extras
for ($k = 1; $k -le $deck.Slides.Count; $k++) {
    $deck.Slides($k).Export((Join-Path $rdir ('VimKD_PI_summary_slide{0:D2}.png' -f $k)), 'PNG', 2000, 1125)
}
"slides=$($deck.Slides.Count) -> $out"
$deck.Close()
# PowerPoint is NOT quit: $app.Quit() would close every deck open on this machine.
