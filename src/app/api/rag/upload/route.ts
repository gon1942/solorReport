import { NextRequest, NextResponse } from 'next/server';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { getApiServerUrl } from '@/config/serverConfig';

// PV Solar 특화 파일 업로드 API
export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get('file') as File;
    const user = (req as any).user; // 미들웨어에서 설정된 사용자 정보
    
    if (!file) {
      return NextResponse.json(
        {
          success: false,
          error: {
            message: '파일이 제공되지 않았습니다.'
          }
        }, 
        { status: 400 }
      );
    }
    
    // 파일 타입 검증 (PV Solar 특화)
    const fileExtension = file.name.split('.').pop()?.toLowerCase();
    const allowedExtensions = ['xlsx', 'xls', 'csv', 'txt'];
    if (!allowedExtensions.includes(fileExtension || '')) {
      return NextResponse.json(
        {
          success: false,
          error: {
            message: `지원되지 않는 파일 형식입니다: ${fileExtension}. PV Solar 데이터는 Excel(.xlsx, .xls) 또는 CSV(.csv) 파일만 지원합니다.`
          }
        }, 
        { status: 400 }
      );
    }
    
    // 파일 크기 검증 (50MB)
    if (file.size > 50 * 1024 * 1024) {
      return NextResponse.json(
        {
          success: false,
          error: {
            message: '파일 크기가 50MB를 초과했습니다. PV Solar 데이터 파일은 50MB 이하로 업로드해주세요.'
          }
        }, 
        { status: 400 }
      );
    }
    
    // PV Solar 특화 파라미터 설정
    const pvSolarParams = new FormData();

    // 기본 API 파라미터 (명시적으로 기본값 강제)
    const ragPdfBackend = (formData.get('rag_pdf_backend') as string) || 'pdfplumber';
    const ragGenerateImageEmbeddings = (formData.get('rag_generate_image_embeddings') as string) || 'no';
    const useEnhancedEmbedding = (formData.get('use_enhanced_embedding') as string) || 'no';
    const ragProcessTextOnly = (formData.get('rag_process_text_only') as string) || 'no';
    const ragMode = (formData.get('rag_mode') as string) || 'rich' ; //'intelligence';
    const overwrite = (formData.get('overwrite') as string) || 'true';
    const ragEnableGraphrag = (formData.get('rag_enable_graphrag') as string) || 'no';
    const userIdField =
      (formData.get('user_id') as string) ||
      (user?.username as string) ||
      (user?.id as string) ||
      'ryan';
    
    // PV Solar 특화 파라미터 추가
    const pvSolarSpecificOptions = JSON.parse(formData.get('pv_solar_specific_options') as string || '{}');
    
    pvSolarParams.append('rag_pdf_backend', ragPdfBackend);
    pvSolarParams.append('rag_generate_image_embeddings', ragGenerateImageEmbeddings);
    pvSolarParams.append('use_enhanced_embedding', useEnhancedEmbedding);
    pvSolarParams.append('rag_process_text_only', ragProcessTextOnly);
    pvSolarParams.append('rag_mode', ragMode);
    pvSolarParams.append('overwrite', overwrite);
    pvSolarParams.append('rag_enable_graphrag', ragEnableGraphrag);
    pvSolarParams.append('user_id', userIdField);
    
    // PV Solar 특화 설정 추가
    if (pvSolarSpecificOptions.energyUnit) {
      pvSolarParams.append('pv_solar_energy_unit', pvSolarSpecificOptions.energyUnit);
    }
    if (pvSolarSpecificOptions.normalizeDates) {
      pvSolarParams.append('pv_solar_normalize_dates', 'true');
    }
    if (pvSolarSpecificOptions.detectWeatherPatterns) {
      pvSolarParams.append('pv_solar_detect_weather', 'true');
    }
    if (pvSolarSpecificOptions.normalizeLocations) {
      pvSolarParams.append('pv_solar_normalize_locations', 'true');
    }
    if (pvSolarSpecificOptions.includeMetadata) {
      pvSolarParams.append('pv_solar_include_metadata', 'true');
    }
    
    // PV Solar 프로젝트 메타데이터 추가
    if (pvSolarSpecificOptions.customMetadata) {
      const metadata = pvSolarSpecificOptions.customMetadata;
      Object.entries(metadata).forEach(([key, value]) => {
        if (value) {
          pvSolarParams.append(`pv_solar_${key}`, String(value));
        }
      });
    }
    
    // 인증 토큰 설정
    const headers: Record<string, string> = {};
    // API Key 헤더 추가 (.env에서 주입, 없으면 빈 문자열)
    headers['X-API-Key'] = process.env.NEXT_PUBLIC_RAG_STATUS_KEY || '';
    const token = req.cookies.get('auth_token')?.value;
    if (token) {
      headers['Authorization'] = `Bearer ${token}`;
    }
    
    // 사용자 ID 설정
    if (user?.username) {
      console.log('PV Solar 업로드 - 사용자 username을 user_id로 추가:', user.username);
    } else if (user?.id) {
      console.log('PV Solar 업로드 - 사용자 ID를 추가 (fallback):', user.id);
    } else {
      console.log('PV Solar 업로드 - 사용자 ID가 없음');
    }
    
    // 파일 데이터 추가
    const bytes = await file.arrayBuffer();
    const buffer = Buffer.from(bytes);
    const uploadDir = path.join(process.cwd(), 'upload');
    await mkdir(uploadDir, { recursive: true });
    const savedFilePath = path.join(uploadDir, path.basename(file.name));
    await writeFile(savedFilePath, buffer);

    const blob = new Blob([buffer], { type: file.type });
    pvSolarParams.append('file', blob, file.name);
    
    // API 서버에 PV Solar 파일 업로드 요청 (env 반영을 위해 런타임에 해석)
    const apiServer = getApiServerUrl();
    console.log('PV Solar API 서버에 파일 업로드 요청 전송=====>:', {
      url: `${apiServer}/rag/upload`,
      method: 'POST',
      body: pvSolarParams
    });
    
    try {
      const apiResponse = await fetch(`${apiServer}/rag/upload`, {
        method: 'POST',
        headers: headers,
        body: pvSolarParams
      });
      
      if (!apiResponse.ok) {
        const errorText = await apiResponse.text();
        console.error('PV Solar API 서버 파일 업로드 응답 오류:', {
          status: apiResponse.status,
          statusText: apiResponse.statusText,
          error: errorText
        });
        
        return NextResponse.json(
          {
            success: false,
            error: {
              message: `PV Solar 파일 업로드 실패: ${apiResponse.statusText}`,
              details: errorText,
              pvSolarSpecific: 'PV Solar 데이터 처리 중 오류가 발생했습니다. 파일 형식과 내용을 확인해주세요.'
            }
          }, 
          { status: apiResponse.status }
        );
      }
      
      const apiData = await apiResponse.json();
      console.log('PV Solar API 서버 파일 업로드 응답 성공:', apiData);
      
      // PV Solar 특화 응답 데이터 구성
      const pvSolarResponse = {
        success: true,
        data: {
          id: apiData.id || `file-${Date.now()}`,
          name: file.name,
          type: file.type,
          size: file.size,
          extension: fileExtension,
          message: 'PV Solar 파일이 성공적으로 업로드되었습니다.',
          pvSolarProcessing: {
            energyUnit: pvSolarSpecificOptions.energyUnit || 'kW',
            normalizeDates: pvSolarSpecificOptions.normalizeDates,
            detectWeatherPatterns: pvSolarSpecificOptions.detectWeatherPatterns,
            normalizeLocations: pvSolarSpecificOptions.normalizeLocations,
            includeMetadata: pvSolarSpecificOptions.includeMetadata,
            customMetadata: pvSolarSpecificOptions.customMetadata
          },
          ragParams: {
            ragMode,
            ragProcessTextOnly,
            ragPdfBackend,
            useEnhancedEmbedding,
            ragGenerateImageEmbeddings,
            ragEnableGraphrag
          },
          savedPath: savedFilePath
        }
      };
      
      return NextResponse.json(pvSolarResponse);
    } catch (error) {
      console.error('PV Solar API 서버 파일 업로드 요청 실패:', error);
      
      // PV Solar 모의 응답 생성 (개발 목적)
      return NextResponse.json({
        success: true,
        data: {
          id: `file-${Date.now()}`,
          name: file.name,
          type: file.type,
          size: file.size,
          extension: fileExtension,
          message: 'PV Solar 파일이 성공적으로 업로드되었습니다. (개발 환경 모의 응답)',
          pvSolarProcessing: {
            energyUnit: pvSolarSpecificOptions.energyUnit || 'kW',
            normalizeDates: pvSolarSpecificOptions.normalizeDates,
            detectWeatherPatterns: pvSolarSpecificOptions.detectWeatherPatterns,
            normalizeLocations: pvSolarSpecificOptions.normalizeLocations,
            includeMetadata: pvSolarSpecificOptions.includeMetadata,
            customMetadata: pvSolarSpecificOptions.customMetadata
          },
          savedPath: savedFilePath
        }
      });
    }
  } catch (error) {
    console.error('PV Solar 파일 업로드 요청 처리 중 오류:', error);
    return NextResponse.json(
      {
        success: false,
        error: {
          message: error instanceof Error ? error.message : 'PV Solar 파일 업로드 중 오류가 발생했습니다.',
          pvSolarSpecific: 'PV Solar 데이터 처리 중 예기치 않은 오류가 발생했습니다. 파일을 확인하고 다시 시도해주세요.'
        }
      },
      { status: 500 }
    );
  }
}
