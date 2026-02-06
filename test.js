function testLoadSealImage() {
  Logger.log('=== 견적서 폴더에서 직인 이미지 로드 테스트 ===');
  
  try {
    // PDF 저장 폴더에서 직인 이미지 파일 찾기
    const rootFolder = DriveApp.getRootFolder();
    const folderIterator = rootFolder.getFoldersByName('견적서');
    
    if (!folderIterator.hasNext()) {
      Logger.log('오류: 견적서 폴더가 없습니다.');
      return;
    }
    
    const folder = folderIterator.next();
    Logger.log('✓ 견적서 폴더 찾음');
    
    // seal.jpeg 파일 찾기
    const files = folder.getFilesByName('seal.jpeg');
    
    if (files.hasNext()) {
      const file = files.next();
      Logger.log('✓ seal.jpeg 찾음');
      Logger.log('파일명: ' + file.getName());
      Logger.log('MimeType: ' + file.getMimeType());
      
      const fileId = file.getId();
      const sealImageUrl = `https://drive.google.com/uc?id=${fileId}&export=download`;
      
      Logger.log('✓ 직인 이미지 로드 완료!');
      Logger.log('FileID: ' + fileId);
      Logger.log('이미지 URL: ' + sealImageUrl);
      
      return sealImageUrl;
    } else {
      Logger.log('오류: 견적서 폴더에서 seal.jpeg를 찾을 수 없습니다.');
      Logger.log('대안: 견적서 폴더에 seal.jpeg 파일을 업로드하세요.');
    }
    
  } catch (e) {
    Logger.log('오류 발생: ' + e.message);
    Logger.log('스택: ' + e.stack);
  }
  
  Logger.log('=== 테스트 완료 ===');
}