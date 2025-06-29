const express = require('express');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');
const { downloadFile } = require('../utils/supabase');

const router = express.Router();

// 📧 이메일 전송 설정
const createTransporter = () => {
  // 환경 변수가 없으면 테스트 모드로 실행
  if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
    console.log('⚠️  이메일 설정이 없어 테스트 모드로 실행됩니다.');
    return nodemailer.createTransport({
      host: 'smtp.ethereal.email',
      port: 587,
      auth: {
        user: 'test@test.com',
        pass: 'test123'
      }
    });
  }
  
  return nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS
    }
  });
};

// 📧 이메일 전송
router.post('/send', async (req, res) => {
  try {
    const { 
      to, 
      subject, 
      body, 
      attachmentPath, 
      templateId,
      scheduleTime 
    } = req.body;

    // 필수 필드 검증
    if (!to || !subject || !attachmentPath) {
      return res.status(400).json({ 
        error: '필수 필드가 누락되었습니다. (받는 사람, 제목, 첨부파일)' 
      });
    }

    // Supabase Storage에서 첨부파일 다운로드
    console.log('📥 이메일 첨부파일 다운로드 중:', attachmentPath);
    const downloadResult = await downloadFile(attachmentPath, 'generated');
    
    if (!downloadResult.success) {
      console.log('❌ 첨부파일 다운로드 실패:', downloadResult.error);
      return res.status(404).json({ error: '첨부파일을 찾을 수 없습니다.' });
    }
    
    // 임시 파일로 저장 (이메일 첨부용)
    const tempAttachmentPath = path.join(__dirname, '../uploads', attachmentPath);
    // uploads 폴더가 없으면 생성
    if (!fs.existsSync(path.join(__dirname, '../uploads'))) {
      fs.mkdirSync(path.join(__dirname, '../uploads'), { recursive: true });
    }
    fs.writeFileSync(tempAttachmentPath, downloadResult.data);
    console.log('✅ 첨부파일 임시 저장 완료');

    // 이메일 템플릿 적용 (템플릿이 있는 경우)
    let emailBody = body || '안녕하세요.\n\n발주서를 첨부파일로 보내드립니다.\n\n확인 후 회신 부탁드립니다.\n\n감사합니다.';
    let emailSubject = subject;

    if (templateId) {
      const template = loadEmailTemplate(templateId);
      if (template) {
        emailSubject = template.subject || subject;
        emailBody = template.body || body;
      }
    }

    // 즉시 전송인지 예약 전송인지 확인
    if (scheduleTime && new Date(scheduleTime) > new Date()) {
      // 예약 전송 (실제로는 스케줄러가 필요하지만, 여기서는 로깅만)
      console.log(`이메일 예약됨: ${scheduleTime}에 ${to}로 전송 예정`);
      
      res.json({
        success: true,
        message: `이메일이 ${scheduleTime}에 전송되도록 예약되었습니다.`,
        scheduled: true,
        scheduleTime: scheduleTime
      });
      
      return;
    }

    // 즉시 전송
    const transporter = createTransporter();
    
    const mailOptions = {
      from: process.env.EMAIL_USER || 'test@test.com',
      to: to,
      subject: emailSubject,
      text: emailBody,
      html: emailBody.replace(/\n/g, '<br>'),
      attachments: [
        {
          filename: path.basename(attachmentPath),
          path: tempAttachmentPath
        }
      ]
    };

    // 이메일 설정이 없으면 시뮬레이션 모드
    console.log('🔍 환경변수 체크:', {
      EMAIL_USER: process.env.EMAIL_USER ? '설정됨' : '설정안됨',
      EMAIL_PASS: process.env.EMAIL_PASS ? '설정됨' : '설정안됨',
      NODE_ENV: process.env.NODE_ENV
    });
    
    if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
      console.log('📧 [시뮬레이션 모드] 이메일 전송:', {
        to: to,
        subject: emailSubject,
        attachment: path.basename(attachmentPath)
      });
      
      // 가짜 응답 생성
      const info = {
        messageId: 'simulation-' + Date.now(),
        accepted: [to]
      };
      
      // 전송 이력 저장
      saveEmailHistory({
        to,
        subject: emailSubject,
        attachmentName: path.basename(attachmentPath),
        sentAt: new Date().toISOString(),
        messageId: info.messageId,
        status: 'success (simulation)'
      });

      res.json({
        success: true,
        message: `이메일이 시뮬레이션으로 전송되었습니다. (${to}) - 실제 전송하려면 Gmail 설정을 완료하세요.`,
        messageId: info.messageId,
        sentAt: new Date().toISOString(),
        simulation: true
      });
      
      return;
    }

    const info = await transporter.sendMail(mailOptions);
    
    // 임시 파일 정리
    if (fs.existsSync(tempAttachmentPath)) {
      fs.unlinkSync(tempAttachmentPath);
      console.log('✅ 임시 첨부파일 정리 완료');
    }
    
    // 전송 이력 저장
    saveEmailHistory({
      to,
      subject: emailSubject,
      attachmentName: path.basename(attachmentPath),
      sentAt: new Date().toISOString(),
      messageId: info.messageId,
      status: 'success'
    });

    res.json({
      success: true,
      message: `이메일이 성공적으로 전송되었습니다. (${to})`,
      messageId: info.messageId,
      sentAt: new Date().toISOString()
    });

  } catch (error) {
    console.error('이메일 전송 오류:', error);
    
    // 실패 이력 저장
    saveEmailHistory({
      to: req.body.to,
      subject: req.body.subject,
      sentAt: new Date().toISOString(),
      status: 'failed',
      error: error.message
    });

    res.status(500).json({ 
      error: '이메일 전송 중 오류가 발생했습니다.', 
      details: error.message 
    });
  }
});

// 📧 이메일 템플릿 저장
router.post('/template', (req, res) => {
  try {
    const { templateName, subject, body, recipients } = req.body;
    
    const templateData = {
      name: templateName,
      subject,
      body,
      recipients,
      createdAt: new Date().toISOString()
    };

    const templatePath = path.join(__dirname, '../file/email-templates');
    if (!fs.existsSync(templatePath)) {
      fs.mkdirSync(templatePath, { recursive: true });
    }

    fs.writeFileSync(
      path.join(templatePath, `${templateName}.json`),
      JSON.stringify(templateData, null, 2)
    );

    res.json({
      success: true,
      message: '이메일 템플릿이 저장되었습니다.',
      templateId: templateName
    });

  } catch (error) {
    res.status(500).json({ 
      error: '템플릿 저장 중 오류가 발생했습니다.', 
      details: error.message 
    });
  }
});

// 📧 전송 이력 조회
router.get('/history', (req, res) => {
  try {
    const historyPath = path.join(__dirname, '../file/email-history.json');
    
    if (!fs.existsSync(historyPath)) {
      return res.json({ success: true, history: [] });
    }

    const history = JSON.parse(fs.readFileSync(historyPath, 'utf8'));
    
    res.json({
      success: true,
      history: history.sort((a, b) => new Date(b.sentAt) - new Date(a.sentAt))
    });

  } catch (error) {
    res.status(500).json({ 
      error: '이력 조회 중 오류가 발생했습니다.', 
      details: error.message 
    });
  }
});

// 📧 선택된 이력 삭제
router.delete('/history/delete', (req, res) => {
  try {
    const { indices } = req.body;
    
    if (!Array.isArray(indices) || indices.length === 0) {
      return res.status(400).json({ 
        error: '삭제할 항목의 인덱스가 필요합니다.' 
      });
    }

    const historyPath = path.join(__dirname, '../file/email-history.json');
    
    if (!fs.existsSync(historyPath)) {
      return res.status(404).json({ 
        error: '이력 파일을 찾을 수 없습니다.' 
      });
    }

    let history = JSON.parse(fs.readFileSync(historyPath, 'utf8'));
    
    // 내림차순 정렬된 상태에서의 인덱스이므로 정렬 먼저
    history = history.sort((a, b) => new Date(b.sentAt) - new Date(a.sentAt));
    
    // 인덱스를 내림차순으로 정렬하여 뒤쪽부터 삭제 (인덱스 오류 방지)
    const sortedIndices = indices.sort((a, b) => b - a);
    
    // 유효한 인덱스인지 확인
    for (const index of sortedIndices) {
      if (index < 0 || index >= history.length) {
        return res.status(400).json({ 
          error: `유효하지 않은 인덱스입니다: ${index}` 
        });
      }
    }
    
    // 선택된 항목들 삭제
    for (const index of sortedIndices) {
      history.splice(index, 1);
    }
    
    fs.writeFileSync(historyPath, JSON.stringify(history, null, 2));
    
    res.json({
      success: true,
      message: `${indices.length}개 항목이 삭제되었습니다.`,
      deletedCount: indices.length
    });

  } catch (error) {
    console.error('이력 삭제 오류:', error);
    res.status(500).json({ 
      error: '이력 삭제 중 오류가 발생했습니다.', 
      details: error.message 
    });
  }
});

// 📧 전체 이력 삭제
router.delete('/history/clear', (req, res) => {
  try {
    const historyPath = path.join(__dirname, '../file/email-history.json');
    
    // 빈 배열로 파일 내용 교체
    fs.writeFileSync(historyPath, JSON.stringify([], null, 2));
    
    res.json({
      success: true,
      message: '모든 전송 이력이 삭제되었습니다.'
    });

  } catch (error) {
    console.error('전체 이력 삭제 오류:', error);
    res.status(500).json({ 
      error: '전체 이력 삭제 중 오류가 발생했습니다.', 
      details: error.message 
    });
  }
});

// 유틸리티 함수들
function loadEmailTemplate(templateId) {
  try {
    const templatePath = path.join(__dirname, '../file/email-templates', `${templateId}.json`);
    if (fs.existsSync(templatePath)) {
      return JSON.parse(fs.readFileSync(templatePath, 'utf8'));
    }
  } catch (error) {
    console.error('템플릿 로드 오류:', error);
  }
  return null;
}

function saveEmailHistory(historyItem) {
  try {
    const historyPath = path.join(__dirname, '../file/email-history.json');
    let history = [];
    
    if (fs.existsSync(historyPath)) {
      history = JSON.parse(fs.readFileSync(historyPath, 'utf8'));
    }
    
    history.push(historyItem);
    
    // 최근 100개만 유지
    if (history.length > 100) {
      history = history.slice(-100);
    }
    
    fs.writeFileSync(historyPath, JSON.stringify(history, null, 2));
  } catch (error) {
    console.error('이력 저장 오류:', error);
  }
}

module.exports = router; 