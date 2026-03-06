// 全局变量
let houseData = []; // 存储Excel数据
let currentHouse = null; // 当前选中的房屋
let currentPhotoIndex = 0; // 当前照片索引

// 图片原始尺寸
const IMAGE_WIDTH = 1868;
const IMAGE_HEIGHT = 1302;

// 支持的图片格式（大小写不敏感）
const SUPPORTED_IMAGE_EXTENSIONS = ['jpg', 'jpeg', 'png'];

// 页面加载完成后自动加载数据
window.onload = function() {
    loadDataFromServer();
    initCoordinateTracking();
};

// 初始化鼠标坐标追踪
function initCoordinateTracking() {
    const svg = document.getElementById('villageSvg');
    const coordDisplay = document.getElementById('coordinateDisplay');
    const coordValue = document.getElementById('coordinateValue');

    // 鼠标进入图幅时显示坐标
    svg.addEventListener('mouseenter', () => {
        coordDisplay.classList.add('visible');
    });

    // 鼠标离开图幅时隐藏坐标
    svg.addEventListener('mouseleave', () => {
        coordDisplay.classList.remove('visible');
    });

    // 鼠标移动时更新坐标
    svg.addEventListener('mousemove', (e) => {
        // 获取SVG相对于视口的位置和尺寸
        const svgRect = svg.getBoundingClientRect();
        
        // 计算鼠标在SVG内的相对位置（像素）
        const x = e.clientX - svgRect.left;
        const y = e.clientY - svgRect.top;
        
        // 考虑SVG的viewBox和实际显示尺寸的缩放比例
        // 实际显示的图像宽度/高度（保持比例后的实际显示尺寸）
        const scaleX = IMAGE_WIDTH / svgRect.width;
        const scaleY = IMAGE_HEIGHT / svgRect.height;
        
        // 由于 preserveAspectRatio="xMidYMid meet"，图片保持比例居中显示
        // 需要计算实际图像在SVG中的偏移量和缩放
        const imageAspect = IMAGE_WIDTH / IMAGE_HEIGHT;
        const svgAspect = svgRect.width / svgRect.height;
        
        let actualImageWidth, actualImageHeight, offsetX, offsetY;
        
        if (svgAspect > imageAspect) {
            // SVG更宽，图像高度填满，左右有黑边
            actualImageHeight = svgRect.height;
            actualImageWidth = actualImageHeight * imageAspect;
            offsetX = (svgRect.width - actualImageWidth) / 2;
            offsetY = 0;
        } else {
            // SVG更高，图像宽度填满，上下有黑边
            actualImageWidth = svgRect.width;
            actualImageHeight = actualImageWidth / imageAspect;
            offsetX = 0;
            offsetY = (svgRect.height - actualImageHeight) / 2;
        }
        
        // 检查鼠标是否在图像范围内（不在黑边区域）
        if (x < offsetX || x > offsetX + actualImageWidth || 
            y < offsetY || y > offsetY + actualImageHeight) {
            // 在黑边区域，不更新坐标或显示特殊提示
            coordValue.textContent = 'X: --.- Y: --.-';
            coordValue.style.color = '#999';
            return;
        }
        
        coordValue.style.color = 'white';
        
        // 计算在图像内的相对位置（0-1范围）
        const relativeX = (x - offsetX) / actualImageWidth;
        const relativeY = (y - offsetY) / actualImageHeight;
        
        // 转换为百分比（保留一位小数）
        const percentX = (relativeX * 100).toFixed(1);
        const percentY = (relativeY * 100).toFixed(1);
        
        // 更新显示
        coordValue.textContent = `X: ${percentX}% Y: ${percentY}%`;
    });
}

// 从服务器自动加载Excel数据
async function loadDataFromServer() {
    try {
        updateStatus('正在连接服务器...', true);
        
        // 使用 Fetch API 获取同目录下的 Information.xlsx
        const response = await fetch('Information.xlsx');
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        // 获取二进制数据
        const arrayBuffer = await response.arrayBuffer();
        
        // 使用 SheetJS 解析
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // 转换为JSON
        houseData = XLSX.utils.sheet_to_json(worksheet);
        
        if (houseData.length === 0) {
            throw new Error('Excel文件中没有数据');
        }
        
        updateStatus(`成功加载 ${houseData.length} 条农房数据`, false);
        
        // 渲染热区
        renderHotspots();
        
        // 自动选中第一个房屋
        selectHouse(houseData[0]);
        
        // 隐藏加载提示
        document.getElementById('loadingStatus').style.display = 'none';
        
    } catch (error) {
        console.error('自动加载失败:', error);
        updateStatus(`自动加载失败: ${error.message}，请检查文件路径`, false);
        
        // 如果自动加载失败，显示错误信息
        document.getElementById('emptyState').innerHTML = `
            <div class="empty-icon">⚠️</div>
            <div>数据加载失败</div>
            <div style="font-size: 14px; margin-top: 8px; opacity: 0.7; color: #e74c3c;">
                ${error.message}<br>
                请确保 Information.xlsx 与网页在同一目录
            </div>
        `;
        document.getElementById('loadingStatus').innerHTML = '❌ 加载失败';
    }
}

// 更新状态栏
function updateStatus(message, loading) {
    const statusText = document.getElementById('statusText');
    const loadingStatus = document.getElementById('loadingStatus');
    
    statusText.textContent = message;
    
    if (loading) {
        loadingStatus.innerHTML = '<span class="loading-spinner"></span>正在加载...';
    } else {
        loadingStatus.style.display = 'none';
    }
}

// 判断坐标格式并标准化为实际像素坐标
function normalizeCoordinate(value, maxDimension) {
    if (value === undefined || value === null) return 0;
    const num = parseFloat(value);
    if (isNaN(num)) return 0;
    
    // 如果值小于等于1，假设是0-1范围的小数（百分比），转换为像素
    if (num > 0 && num <= 1) {
        return num * maxDimension;
    }
    // 如果值小于等于100，假设是0-100范围的百分比，转换为像素
    if (num > 1 && num <= 100) {
        return (num / 100) * maxDimension;
    }
    // 否则假设已经是实际像素坐标
    return num;
}

// 渲染SVG热区
function renderHotspots() {
    const group = document.getElementById('hotspotsGroup');
    const labelsGroup = document.getElementById('labelsGroup');
    group.innerHTML = '';
    labelsGroup.innerHTML = '';

    houseData.forEach((house, index) => {
        // 获取并标准化坐标（支持百分比、小数或像素）
        const x1 = normalizeCoordinate(house.X1, IMAGE_WIDTH);
        const y1 = normalizeCoordinate(house.Y1, IMAGE_HEIGHT);
        const x2 = normalizeCoordinate(house.X2, IMAGE_WIDTH);
        const y2 = normalizeCoordinate(house.Y2, IMAGE_HEIGHT);
        const x3 = normalizeCoordinate(house.X3, IMAGE_WIDTH);
        const y3 = normalizeCoordinate(house.Y3, IMAGE_HEIGHT);
        const x4 = normalizeCoordinate(house.X4, IMAGE_WIDTH);
        const y4 = normalizeCoordinate(house.Y4, IMAGE_HEIGHT);

        // 检查坐标数据是否有效
        if (x1 === 0 && y1 === 0 && x2 === 0 && y2 === 0) {
            console.warn('房屋数据缺少有效坐标信息:', house);
            return;
        }

        // 创建四边形热区（使用实际像素坐标）
        const polygon = document.createElementNS('http://www.w3.org/2000/svg', 'polygon');
        const points = `${x1},${y1} ${x2},${y2} ${x3},${y3} ${x4},${y4}`;
        polygon.setAttribute('points', points);
        polygon.setAttribute('class', 'hotspot');
        polygon.setAttribute('data-index', index);
        polygon.setAttribute('data-code', house.房屋编码);
        
        // 添加事件监听
        polygon.addEventListener('click', () => selectHouse(house));
        polygon.addEventListener('mouseenter', (e) => showTooltip(e, house.房屋名称 || house.房屋编码));
        polygon.addEventListener('mouseleave', hideTooltip);
        
        group.appendChild(polygon);

        // 添加文字标签（显示房屋编码）
        const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        const centerX = (x1 + x2 + x3 + x4) / 4;
        const centerY = (y1 + y2 + y3 + y4) / 4;
        text.setAttribute('x', centerX);
        text.setAttribute('y', centerY);
        text.setAttribute('class', 'hotspot-label');
        text.textContent = house.房屋编码;
        text.style.pointerEvents = 'none';
        labelsGroup.appendChild(text);
    });
}

// 选中房屋
function selectHouse(house) {
    currentHouse = house;
    currentPhotoIndex = 0;

    // 更新热区样式
    document.querySelectorAll('.hotspot').forEach(h => {
        h.classList.remove('active');
        if (h.getAttribute('data-code') === house.房屋编码) {
            h.classList.add('active');
        }
    });

    // 显示信息面板
    document.getElementById('emptyState').style.display = 'none';
    document.getElementById('houseInfo').style.display = 'block';

    // 填充基本信息
    document.getElementById('houseCode').textContent = house.房屋编码 || '-';
    document.getElementById('houseName').textContent = house.房屋名称 || '-';
    document.getElementById('buildYear').textContent = house.建成年代 || '-';
    document.getElementById('area').textContent = house.占地面积 || '-';

    // 加载照片（自动识别多种格式）
    loadPhotos(house.房屋编码);
}

// 尝试加载图片，支持多种格式（大小写不敏感）
async function tryLoadImage(basePath) {
    // 生成所有可能的格式组合（原始大小写 + 全小写 + 全大写）
    const variations = [];
    
    SUPPORTED_IMAGE_EXTENSIONS.forEach(ext => {
        // 小写扩展名
        variations.push(`${basePath}.${ext.toLowerCase()}`);
        // 大写扩展名
        variations.push(`${basePath}.${ext.toUpperCase()}`);
        // 首字母大写
        variations.push(`${basePath}.${ext.charAt(0).toUpperCase() + ext.slice(1).toLowerCase()}`);
    });
    
    // 去重
    const uniqueVariations = [...new Set(variations)];
    
    // 尝试加载每个可能的URL
    for (const url of uniqueVariations) {
        try {
            const img = new Image();
            const loadPromise = new Promise((resolve, reject) => {
                img.onload = () => resolve({ success: true, url: url, img: img });
                img.onerror = () => reject({ success: false, url: url });
            });
            
            img.src = url;
            const result = await loadPromise;
            return result;
        } catch (e) {
            // 继续尝试下一个格式
            continue;
        }
    }
    
    // 所有格式都失败
    return { success: false, url: null, img: null };
}

// 加载照片（改进版：自动识别 jpg/jpeg/png，大小写不敏感）
function loadPhotos(houseCode) {
    const maxPhotos = 3;
    
    for (let i = 0; i < maxPhotos; i++) {
        const slide = document.getElementById(`slide${i}`);
        const photoBasePath = `photo/${houseCode}-${i + 1}`;
        
        slide.innerHTML = '<div class="photo-placeholder">加载中...</div>';
        
        // 使用异步函数尝试加载图片
        (async (slideElement) => {
            const result = await tryLoadImage(photoBasePath);
            
            if (result.success) {
                slideElement.innerHTML = '';
                result.img.style.maxWidth = '100%';
                result.img.style.maxHeight = '100%';
                result.img.style.objectFit = 'contain';
                slideElement.appendChild(result.img);
                console.log(`成功加载图片: ${result.url}`);
            } else {
                slideElement.innerHTML = '<div class="photo-placeholder">暂无照片</div>';
            }
            
            updatePhotoControls();
        })(slide);
    }
    
    updatePhotoControls();
}

// 更新照片控制状态
function updatePhotoControls() {
    const slides = document.querySelectorAll('.photo-slide');
    const indicators = document.querySelectorAll('.indicator');
    
    slides.forEach((slide, index) => {
        slide.classList.toggle('active', index === currentPhotoIndex);
    });
    
    indicators.forEach((ind, index) => {
        ind.classList.toggle('active', index === currentPhotoIndex);
    });

    document.getElementById('photoCounter').textContent = 
        `${currentPhotoIndex + 1} / ${slides.length}`;
}

// 切换照片
function changePhoto(direction) {
    const slides = document.querySelectorAll('.photo-slide');
    let newIndex = currentPhotoIndex + direction;
    
    if (newIndex < 0) newIndex = slides.length - 1;
    if (newIndex >= slides.length) newIndex = 0;
    
    currentPhotoIndex = newIndex;
    updatePhotoControls();
}

// 跳转到指定照片
function goToPhoto(index) {
    currentPhotoIndex = index;
    updatePhotoControls();
}

// 显示提示框
function showTooltip(e, text) {
    const tooltip = document.getElementById('tooltip');
    tooltip.textContent = text;
    tooltip.style.display = 'block';
    
    const rect = e.target.getBoundingClientRect();
    const containerRect = document.getElementById('svgContainer').getBoundingClientRect();
    
    tooltip.style.left = (rect.left - containerRect.left + rect.width / 2) + 'px';
    tooltip.style.top = (rect.top - containerRect.top - 30) + 'px';
    tooltip.style.transform = 'translateX(-50%)';
}

// 隐藏提示框
function hideTooltip() {
    document.getElementById('tooltip').style.display = 'none';
}

// 处理上传按钮点击
function handleUpload() {
    if (!currentHouse) {
        alert('请先选择一个农房');
        return;
    }
    alert(`准备为 ${currentHouse.房屋编码} - ${currentHouse.房屋名称} 上传资料\n\n实际项目中这里会打开文件选择对话框`);
}

// 键盘支持
document.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowLeft') changePhoto(-1);
    if (e.key === 'ArrowRight') changePhoto(1);
});