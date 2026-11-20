(() => {
  const MODEL_URL = '/static/bg-remover/u2netp.onnx';
  const INPUT_SIZE = 320;
  const MAX_SIDE = 768;
  let sessionPromise = null;

  async function getSession() {
    if (!window.ort) {
      throw new Error('ONNX Runtime Web failed to load.');
    }
    if (!sessionPromise) {
      if (window.ort.env && window.ort.env.wasm) {
        window.ort.env.wasm.wasmPaths = {
          mjs: '/static/bg-remover/ort-wasm-simd-threaded.mjs',
          wasm: '/static/bg-remover/ort-wasm-simd-threaded.wasm'
        };
        window.ort.env.wasm.numThreads = 1;
        window.ort.env.wasm.simd = false;
        window.ort.env.wasm.proxy = false;
      }
      sessionPromise = window.ort.InferenceSession.create(MODEL_URL, {
        executionProviders: ['wasm']
      });
    }
    return sessionPromise;
  }

  function loadImage(file) {
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(file);
      const img = new Image();
      img.onload = () => {
        URL.revokeObjectURL(url);
        resolve(img);
      };
      img.onerror = () => {
        URL.revokeObjectURL(url);
        reject(new Error('Failed to load the selected image.'));
      };
      img.src = url;
    });
  }

  function prepareCanvas(img) {
    let { width, height } = img;
    if (width > MAX_SIDE || height > MAX_SIDE) {
      if (width >= height) {
        height = Math.round((height * MAX_SIDE) / width);
        width = MAX_SIDE;
      } else {
        width = Math.round((width * MAX_SIDE) / height);
        height = MAX_SIDE;
      }
    }
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0, width, height);
    return { canvas, ctx, width, height };
  }

  function toTensor(sourceCanvas) {
    const tmp = document.createElement('canvas');
    tmp.width = INPUT_SIZE;
    tmp.height = INPUT_SIZE;
    const tctx = tmp.getContext('2d');
    tctx.drawImage(sourceCanvas, 0, 0, INPUT_SIZE, INPUT_SIZE);
    const { data } = tctx.getImageData(0, 0, INPUT_SIZE, INPUT_SIZE);

    const area = INPUT_SIZE * INPUT_SIZE;
    const input = new Float32Array(3 * area);
    const mean = [0.485, 0.456, 0.406];
    const std = [0.229, 0.224, 0.225];

    for (let i = 0; i < area; i++) {
      input[i] = (data[i * 4] / 255 - mean[0]) / std[0];
      input[area + i] = (data[i * 4 + 1] / 255 - mean[1]) / std[1];
      input[2 * area + i] = (data[i * 4 + 2] / 255 - mean[2]) / std[2];
    }

    return new window.ort.Tensor('float32', input, [1, 3, INPUT_SIZE, INPUT_SIZE]);
  }

  function applyMask(work, maskTensor) {
    const dims = maskTensor.dims;
    const mh = dims[dims.length - 2];
    const mw = dims[dims.length - 1];
    const raw = maskTensor.data;

    let min = Infinity;
    let max = -Infinity;
    for (let i = 0; i < raw.length; i++) {
      if (raw[i] < min) min = raw[i];
      if (raw[i] > max) max = raw[i];
    }
    const diff = max - min || 1;

    const maskCanvas = document.createElement('canvas');
    maskCanvas.width = mw;
    maskCanvas.height = mh;
    const mctx = maskCanvas.getContext('2d');
    const imgData = mctx.createImageData(mw, mh);

    for (let i = 0; i < mw * mh; i++) {
      const alpha = Math.round(Math.max(0, Math.min(1, (raw[i] - min) / diff)) * 255);
      imgData.data[i * 4] = 255;
      imgData.data[i * 4 + 1] = 255;
      imgData.data[i * 4 + 2] = 255;
      imgData.data[i * 4 + 3] = alpha;
    }
    mctx.putImageData(imgData, 0, 0);

    const fullMask = document.createElement('canvas');
    fullMask.width = work.width;
    fullMask.height = work.height;
    const fctx = fullMask.getContext('2d');
    fctx.drawImage(maskCanvas, 0, 0, work.width, work.height);
    const alphaData = fctx.getImageData(0, 0, work.width, work.height).data;

    const target = work.ctx.getImageData(0, 0, work.width, work.height);
    for (let i = 0; i < work.width * work.height; i++) {
      target.data[i * 4 + 3] = alphaData[i * 4 + 3];
    }
    work.ctx.putImageData(target, 0, 0);
  }

  window.FileVellumBgRemover = {
    async removeBackground(file) {
      const session = await getSession();
      const img = await loadImage(file);
      const work = prepareCanvas(img);
      const tensor = toTensor(work.canvas);

      const output = await session.run({ [session.inputNames[0]]: tensor });
      applyMask(work, output[session.outputNames[0]]);

      return new Promise((resolve, reject) => {
        work.canvas.toBlob((blob) => {
          if (!blob) return reject(new Error('Failed to generate PNG output.'));
          resolve(blob);
        }, 'image/png');
      });
    }
  };
})();
