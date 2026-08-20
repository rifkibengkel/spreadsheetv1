module.exports = {
  apps: [
    {
      name: "sheet-V-2",
      cwd: __dirname,
      script: "node_modules/next/dist/bin/next",
      args: "start",
      interpreter: "node",
      autorestart: true,
      watch: false,
      max_memory_restart: "2G"
    }
  ]
};