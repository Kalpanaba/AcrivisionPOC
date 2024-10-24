import React, { useEffect, useState } from 'react';
import * as XLSX from 'xlsx';
import * as THREE from 'three';
import { OrbitControls } from 'three/examples/jsm/controls/OrbitControls';

const ExcelThreeJSVisualizer = () => {
  const [AData, setAData] = useState([]);
  const [BData, setBData] = useState([]);
  const [CData, setCData] = useState([]);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      if (!workbook.Sheets['A'] || !workbook.Sheets['B'] || !workbook.Sheets['C']) {
        alert('Required sheets (A, B, C) not found');
        return;
      }

      const A = XLSX.utils.sheet_to_json(workbook.Sheets['A']);
      const B = XLSX.utils.sheet_to_json(workbook.Sheets['B']);
      const C = XLSX.utils.sheet_to_json(workbook.Sheets['C']);

      setAData(A);
      setBData(B);
      setCData(C);
    };
    reader.readAsArrayBuffer(file);
  };

  const initializeThreeJS = (A, B, C) => {
    // Scene setup
    const scene = new THREE.Scene();
    scene.background = new THREE.Color(0xffffff);  // White background
    
    // Camera setup
    const camera = new THREE.PerspectiveCamera(
      60,
      window.innerWidth / window.innerHeight,
      0.1,
      1000
    );
    camera.position.set(100, 100, 100);
    
    // Renderer setup
    const renderer = new THREE.WebGLRenderer({ 
      antialias: true,
      alpha: true 
    });
    renderer.setSize(window.innerWidth, window.innerHeight);
    renderer.setPixelRatio(window.devicePixelRatio);
    
    const container = document.getElementById('threejs-container');
    container.innerHTML = '';
    container.appendChild(renderer.domElement);

    // Lighting
    const ambientLight = new THREE.AmbientLight(0xffffff, 0.8);
    scene.add(ambientLight);
    
    const directionalLight = new THREE.DirectionalLight(0xffffff, 0.5);
    directionalLight.position.set(1, 1, 1);
    scene.add(directionalLight);

    // Custom Axes Helper with thicker lines and clear colors
    const createCustomAxis = (length) => {
      const customAxes = new THREE.Group();
      
      // X Axis (Red)
      const xGeometry = new THREE.CylinderGeometry(0.1, 0.1, length, 8);
      const xMaterial = new THREE.MeshBasicMaterial({ color: 0xff0000 });
      const xAxis = new THREE.Mesh(xGeometry, xMaterial);
      xAxis.rotation.z = -Math.PI / 2;
      xAxis.position.x = length / 2;
      
      // Y Axis (Green)
      const yGeometry = new THREE.CylinderGeometry(0.1, 0.1, length, 8);
      const yMaterial = new THREE.MeshBasicMaterial({ color: 0x00ff00 });
      const yAxis = new THREE.Mesh(yGeometry, yMaterial);
      yAxis.position.y = length / 2;
      
      // Z Axis (Blue)
      const zGeometry = new THREE.CylinderGeometry(0.1, 0.1, length, 8);
      const zMaterial = new THREE.MeshBasicMaterial({ color: 0x0000ff });
      const zAxis = new THREE.Mesh(zGeometry, zMaterial);
      zAxis.rotation.x = Math.PI / 2;
      zAxis.position.z = length / 2;
      
      customAxes.add(xAxis);
      customAxes.add(yAxis);
      customAxes.add(zAxis);
      
      return customAxes;
    };

    // Add custom axes
    const customAxes = createCustomAxis(50);
    scene.add(customAxes);

    // Create orbit controls
    const controls = new OrbitControls(camera, renderer.domElement);
    controls.enableDamping = true;
    controls.dampingFactor = 0.05;

    // Create node map
    const nodeMap = {};
    B.forEach(row => {
      nodeMap[row.Node] = {
        x: parseFloat(row.X),
        y: parseFloat(row.Y),
        z: parseFloat(row.Z)
      };
    });

    // Create structural members with cylindrical lines
    A.forEach(({ "Start Node": startNode, "End Node": endNode }) => {
      const start = nodeMap[startNode];
      const end = nodeMap[endNode];

      if (start && end) {
        // Calculate beam direction and length
        const direction = new THREE.Vector3(
          end.x - start.x,
          end.y - start.y,
          end.z - start.z
        );
        const length = direction.length();
        
        // Create cylindrical beam
        const beamGeometry = new THREE.CylinderGeometry(0.9, 0.9, length, 8);
        const beamMaterial = new THREE.MeshPhongMaterial({
          color: 0x156289,
          shininess: 100,
          specular: 0x111111
        });
        
        const beam = new THREE.Mesh(beamGeometry, beamMaterial);
        
        // Position beam
        beam.position.set(
          (start.x + end.x) / 2,
          (start.y + end.y) / 2,
          (start.z + end.z) / 2
        );
        
        // Orient beam
        beam.lookAt(new THREE.Vector3(end.x, end.y, end.z));
        beam.rotateX(Math.PI / 2);
        
        scene.add(beam);
      }
    });

    // Add minimal support nodes for fixed points
    const supportMaterial = new THREE.MeshPhongMaterial({ 
      color: 0x404040,
      transparent: true,
      opacity: 0.8 
    });
    
    C.forEach(({ NodeID, SupportType }) => {
      const node = nodeMap[NodeID];
      if (node && SupportType === 'FIXED') {
        // Create minimal support indicator
        const supportGeometry = new THREE.SphereGeometry(0.3, 16, 16);
        const support = new THREE.Mesh(supportGeometry, supportMaterial);
        support.position.set(node.x, node.y, node.z);
        scene.add(support);
      }
    });

    // Add minimal joint nodes
    const jointMaterial = new THREE.MeshPhongMaterial({ 
      color: 0x000000,
      transparent: true,
      opacity: 0.6 
    });
    
    Object.values(nodeMap).forEach(node => {
      const jointGeometry = new THREE.SphereGeometry(0.25, 16, 16);
      const joint = new THREE.Mesh(jointGeometry, jointMaterial);
      joint.position.set(node.x, node.y, node.z);
      scene.add(joint);
    });

    // Add three cubes representing X, Y, Z axes
    const cubeSize = 2 * 3;  // Increase the size 5 times
    const cubeGeometry = new THREE.BoxGeometry(cubeSize, cubeSize, cubeSize);
    const cubeMaterial = new THREE.MeshPhongMaterial({ color: 0x00ff00 });

    // Cube along the X axis
    const xCube = new THREE.Mesh(cubeGeometry, cubeMaterial);
    xCube.position.set(10, 0, 0);  // Positioned on the X axis
    scene.add(xCube);

    // Cube along the Y axis
    const yCube = new THREE.Mesh(cubeGeometry, cubeMaterial);
    yCube.position.set(0, 10, 0);  // Positioned on the Y axis
    scene.add(yCube);

    // Cube along the Z axis
    const zCube = new THREE.Mesh(cubeGeometry, cubeMaterial);
    zCube.position.set(0, 0, 10);  // Positioned on the Z axis
    scene.add(zCube);

    // Handle window resize
    const handleResize = () => {
      camera.aspect = window.innerWidth / window.innerHeight;
      camera.updateProjectionMatrix();
      renderer.setSize(window.innerWidth, window.innerHeight);
    };
    
    window.addEventListener('resize', handleResize);

    // Animation loop
    const animate = () => {
      requestAnimationFrame(animate);
      controls.update();
      renderer.render(scene, camera);
    };

    animate();

    // Cleanup
    return () => {
      window.removeEventListener('resize', handleResize);
      container.removeChild(renderer.domElement);
    };
  };

  useEffect(() => {
    if (AData.length > 0 && BData.length > 0 && CData.length > 0) {
      initializeThreeJS(AData, BData, CData);
    }
  }, [AData, BData, CData]);

  return (
    <div style={styles.pageContainer}>
      <header style={styles.header}>
        <h1>Excel to 3D Visualizer</h1>
      </header>
      <div style={styles.uploadContainer}>
        <label htmlFor="file-upload" style={styles.label}>
          Upload Excel File
        </label>
        <input
          type="file"
          id="file-upload"
          onChange={handleFileUpload}
          accept=".xlsx, .xls"
          style={styles.input}
        />
      </div>
      <div id="threejs-container" style={styles.visualizerContainer}></div>
    </div>
  );
};

const styles = {
  pageContainer: {
    fontFamily: 'Arial, sans-serif',
    backgroundColor: '#f0f4f8',
    padding: '20px',
    minHeight: '100vh',
    textAlign: 'center',
  },
  header: {
    backgroundColor: '#4682b4',
    color: '#fff',
    padding: '20px',
    borderRadius: '8px',
    marginBottom: '20px',
  },
  uploadContainer: {
    marginBottom: '20px',
  },
  label: {
    fontSize: '18px',
    fontWeight: 'bold',
    marginRight: '10px',
  },
  input: {
    padding: '10px',
    fontSize: '16px',
    borderRadius: '5px',
    border: '1px solid #ccc',
  },
  visualizerContainer: {
    width: '100%',
    height: '500px',
    border: '1px solid #ccc',
    borderRadius: '8px',
    backgroundColor: '#fff',
  },
};

export default ExcelThreeJSVisualizer;
