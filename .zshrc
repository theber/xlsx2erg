workout() {
  aft-mtp-mount ~/mnt
  rsync -av -P /home/tk/Documents/Cycling/training_plan/coverter/*.erg ~/mnt/USB\ storage/plans/
  umount ~/mnt
}
