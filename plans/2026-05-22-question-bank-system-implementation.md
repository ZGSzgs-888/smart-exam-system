# 智能题库与自适应组卷系统 — 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建一个前后端分离的智能题库与自适应组卷系统，支持三种角色（管理员/教师/学生），实现题库管理、智能组卷、在线考试、成绩分析与自适应测评。

**Architecture:** 前端 Vue 3 + Element Plus，后端 Spring Boot + MyBatis-Plus，MySQL 数据库，RESTful API + JWT 认证。组卷引擎采用三层算法架构（约束过滤→均衡分配→能力适配），学生能力模型基于递推公式持续更新。

**Tech Stack:** Vue 3, Element Plus, Vite, Spring Boot, MyBatis-Plus, MySQL, JWT, Apache POI (Excel导入), ECharts (图表)

---

## 项目文件结构

```
exam-system-backend/
├── pom.xml
├── src/main/java/com/exam/
│   ├── ExamApplication.java
│   ├── common/
│   │   ├── config/    WebMvcConfig, SwaggerConfig, CorsConfig
│   │   ├── constant/  RoleConstant, CodeConstant
│   │   ├── exception/ GlobalExceptionHandler, BusinessException
│   │   ├── result/    ApiResult, ResultCode
│   │   └── util/      JwtUtil
│   ├── interceptor/   JwtInterceptor, RoleInterceptor
│   └── module/
│       ├── auth/      登录认证模块
│       ├── user/      用户管理模块
│       ├── subject/   科目管理模块
│       ├── knowledge/ 知识点管理模块
│       ├── question/  题库管理模块
│       ├── paper/     组卷模块 (含 engine/ 算法包)
│       ├── exam/      考试模块
│       ├── score/     成绩模块
│       └── statistics/ 统计模块
├── src/main/resources/
│   ├── application.yml
│   ├── application-dev.yml
│   └── mapper/*.xml
└── sql/init.sql

exam-system-frontend/
├── package.json
├── vite.config.js
├── index.html
├── src/
│   ├── api/           axios 封装 + 各模块 API
│   ├── router/        Vue Router
│   ├── stores/        Pinia 状态管理 (user store)
│   ├── layout/        登录布局 + 管理布局
│   ├── views/
│   │   ├── login/
│   │   ├── admin/     用户管理, 科目管理
│   │   ├── teacher/   题库, 知识点, 组卷, 考试, 成绩, 统计
│   │   └── student/   考试, 成绩, 错题, 分析
│   ├── components/    通用组件 (富文本编辑器, 图表, 上传等)
│   └── utils/         token, 权限, 格式化等工具
```

---

## Phase 1: 项目基础搭建

### Task 1.1: 后端项目脚手架

**Files:**
- Create: `exam-system-backend/pom.xml`
- Create: `exam-system-backend/src/main/java/com/exam/ExamApplication.java`
- Create: `exam-system-backend/src/main/resources/application.yml`
- Create: `exam-system-backend/src/main/resources/application-dev.yml`

- [ ] **Step 1: 创建 pom.xml**

```xml
<?xml version="1.0" encoding="UTF-8"?>
<project xmlns="http://maven.apache.org/POM/4.0.0"
         xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
         xsi:schemaLocation="http://maven.apache.org/POM/4.0.0 http://maven.apache.org/xsd/maven-4.0.0.xsd">
    <modelVersion>4.0.0</modelVersion>
    <parent>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-parent</artifactId>
        <version>3.2.0</version>
    </parent>
    <groupId>com.exam</groupId>
    <artifactId>exam-system</artifactId>
    <version>1.0.0</version>
    <name>智能题库与自适应组卷系统</name>

    <properties>
        <java.version>17</java.version>
    </properties>

    <dependencies>
        <!-- Web -->
        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-web</artifactId>
        </dependency>
        <!-- MyBatis-Plus -->
        <dependency>
            <groupId>com.baomidou</groupId>
            <artifactId>mybatis-plus-spring-boot3-starter</artifactId>
            <version>3.5.6</version>
        </dependency>
        <!-- MySQL -->
        <dependency>
            <groupId>com.mysql</groupId>
            <artifactId>mysql-connector-j</artifactId>
            <scope>runtime</scope>
        </dependency>
        <!-- JWT -->
        <dependency>
            <groupId>com.auth0</groupId>
            <artifactId>java-jwt</artifactId>
            <version>4.4.0</version>
        </dependency>
        <!-- Lombok -->
        <dependency>
            <groupId>org.projectlombok</groupId>
            <artifactId>lombok</artifactId>
            <optional>true</optional>
        </dependency>
        <!-- Validation -->
        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-validation</artifactId>
        </dependency>
        <!-- POI for Excel -->
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>5.2.5</version>
        </dependency>
        <!-- Swagger -->
        <dependency>
            <groupId>com.github.xiaoymin</groupId>
            <artifactId>knife4j-openapi3-jakarta-spring-boot-starter</artifactId>
            <version>4.5.0</version>
        </dependency>
        <!-- Test -->
        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-test</artifactId>
            <scope>test</scope>
        </dependency>
    </dependencies>

    <build>
        <plugins>
            <plugin>
                <groupId>org.springframework.boot</groupId>
                <artifactId>spring-boot-maven-plugin</artifactId>
                <configuration>
                    <excludes>
                        <exclude>
                            <groupId>org.projectlombok</groupId>
                            <artifactId>lombok</artifactId>
                        </exclude>
                    </excludes>
                </configuration>
            </plugin>
        </plugins>
    </build>
</project>
```

- [ ] **Step 2: 创建 Spring Boot 启动类与配置文件**

```java
// ExamApplication.java
package com.exam;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExamApplication {
    public static void main(String[] args) {
        SpringApplication.run(ExamApplication.class, args);
    }
}
```

```yaml
# application.yml
server:
  port: 8080

spring:
  profiles:
    active: dev
  jackson:
    date-format: yyyy-MM-dd HH:mm:ss
    time-zone: Asia/Shanghai

mybatis-plus:
  configuration:
    map-underscore-to-camel-case: true
    log-impl: org.apache.ibatis.logging.stdout.StdOutImpl
  global-config:
    db-config:
      id-type: auto
      logic-delete-field: deleted
      logic-delete-value: 1
      logic-not-delete-value: 0
```

```yaml
# application-dev.yml
spring:
  datasource:
    driver-class-name: com.mysql.cj.jdbc.Driver
    url: jdbc:mysql://localhost:3306/exam_system?useUnicode=true&characterEncoding=utf-8&serverTimezone=Asia/Shanghai
    username: root
    password: 123456
```

- [ ] **Step 3: 创建公共结果类与异常处理**

```java
// ApiResult.java
package com.exam.common.result;

import lombok.Data;

@Data
public class ApiResult<T> {
    private int code;
    private String message;
    private T data;

    public static <T> ApiResult<T> success(T data) {
        ApiResult<T> result = new ApiResult<>();
        result.code = 200;
        result.message = "success";
        result.data = data;
        return result;
    }

    public static <T> ApiResult<T> error(int code, String message) {
        ApiResult<T> result = new ApiResult<>();
        result.code = code;
        result.message = message;
        return result;
    }
}
```

```java
// BusinessException.java
package com.exam.common.exception;

public class BusinessException extends RuntimeException {
    private final int code;
    public BusinessException(int code, String message) {
        super(message);
        this.code = code;
    }
    public int getCode() { return code; }
}
```

```java
// GlobalExceptionHandler.java
package com.exam.common.exception;

import com.exam.common.result.ApiResult;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;

@RestControllerAdvice
public class GlobalExceptionHandler {
    @ExceptionHandler(BusinessException.class)
    public ApiResult<?> handleBusiness(BusinessException e) {
        return ApiResult.error(e.getCode(), e.getMessage());
    }
}
```

- [ ] **Step 4: 创建跨域配置**

```java
package com.exam.common.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.cors.CorsConfiguration;
import org.springframework.web.cors.UrlBasedCorsConfigurationSource;
import org.springframework.web.filter.CorsFilter;

@Configuration
public class CorsConfig {
    @Bean
    public CorsFilter corsFilter() {
        CorsConfiguration config = new CorsConfiguration();
        config.addAllowedOriginPattern("*");
        config.addAllowedHeader("*");
        config.addAllowedMethod("*");
        config.setAllowCredentials(true);
        UrlBasedCorsConfigurationSource source = new UrlBasedCorsConfigurationSource();
        source.registerCorsConfiguration("/**", config);
        return new CorsFilter(source);
    }
}
```

- [ ] **Step 5: 验证后端启动**

Run: 在 `exam-system-backend/` 目录执行 `mvn spring-boot:run`，确认无报错正常启动。

- [ ] **Step 6: 提交**

```bash
git add exam-system-backend/
git commit -m "feat: init backend project scaffold with Spring Boot"
```

---

### Task 1.2: 前端项目脚手架

**Files:**
- Create: `exam-system-frontend/package.json`
- Create: `exam-system-frontend/vite.config.js`
- Create: `exam-system-frontend/index.html`
- Create: `exam-system-frontend/src/main.js`
- Create: `exam-system-frontend/src/App.vue`

- [ ] **Step 1: 初始化 Vue 3 项目**

```bash
cd exam-system-frontend
npm init -y
npm install vue@3 vue-router@4 pinia axios element-plus @element-plus/icons-vue echarts vue-echarts
npm install -D vite @vitejs/plugin-vue
```

- [ ] **Step 2: 创建 vite.config.js**

```javascript
import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'

export default defineConfig({
  plugins: [vue()],
  server: {
    port: 3000,
    proxy: {
      '/api': {
        target: 'http://localhost:8080',
        changeOrigin: true
      }
    }
  },
  resolve: {
    alias: {
      '@': '/src'
    }
  }
})
```

- [ ] **Step 3: 创建 index.html 和 main.js**

```html
<!-- index.html -->
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>智能题库与自适应组卷系统</title>
</head>
<body>
  <div id="app"></div>
  <script type="module" src="/src/main.js"></script>
</body>
</html>
```

```javascript
// src/main.js
import { createApp } from 'vue'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import * as ElementPlusIconsVue from '@element-plus/icons-vue'
import App from './App.vue'
import router from './router'
import { createPinia } from 'pinia'

const app = createApp(App)
app.use(ElementPlus)
app.use(router)
app.use(createPinia())
for (const [key, component] of Object.entries(ElementPlusIconsVue)) {
  app.component(key, component)
}
app.mount('#app')
```

```vue
<!-- src/App.vue -->
<template>
  <router-view />
</template>
```

- [ ] **Step 4: 创建 axios 封装**

```javascript
// src/api/request.js
import axios from 'axios'
import { ElMessage } from 'element-plus'
import router from '@/router'

const request = axios.create({
  baseURL: '/api',
  timeout: 15000
})

request.interceptors.request.use(config => {
  const token = localStorage.getItem('token')
  if (token) {
    config.headers.Authorization = `Bearer ${token}`
  }
  return config
})

request.interceptors.response.use(
  response => {
    const res = response.data
    if (res.code !== 200) {
      ElMessage.error(res.message || '请求失败')
      if (res.code === 401) {
        localStorage.removeItem('token')
        router.push('/login')
      }
      return Promise.reject(new Error(res.message))
    }
    return res
  },
  error => {
    ElMessage.error(error.message || '网络错误')
    return Promise.reject(error)
  }
)

export default request
```

- [ ] **Step 5: 创建路由基础文件**

```javascript
// src/router/index.js
import { createRouter, createWebHistory } from 'vue-router'

const routes = [
  {
    path: '/login',
    name: 'Login',
    component: () => import('@/views/login/LoginView.vue')
  }
]

const router = createRouter({
  history: createWebHistory(),
  routes
})

export default router
```

- [ ] **Step 6: 验证前端启动**

Run: `cd exam-system-frontend && npx vite`，确认在浏览器可访问 localhost:3000。

- [ ] **Step 7: 提交**

```bash
git add exam-system-frontend/
git commit -m "feat: init frontend project scaffold with Vue 3 + Element Plus"
```

---

## Phase 2: 数据库与 JWT 认证

### Task 2.1: 数据库初始化脚本 + JWT 工具

**Files:**
- Modify: `exam-system-backend/src/main/resources/application.yml` (加 JWT 配置)
- Create: `exam-system-backend/sql/init.sql` (完整建表+种子数据)
- Create: `exam-system-backend/src/main/java/com/exam/common/util/JwtUtil.java`

- [ ] **Step 1: 在 application.yml 中添加 JWT 配置**

```yaml
# 添加到 application.yml
jwt:
  secret: exam-system-secret-key-2024
  expiration: 86400000  # 24小时
```

- [ ] **Step 2: 创建 JwtUtil**

```java
package com.exam.common.util;

import com.auth0.jwt.JWT;
import com.auth0.jwt.algorithms.Algorithm;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.util.Date;

@Component
public class JwtUtil {
    @Value("${jwt.secret}")
    private String secret;
    @Value("${jwt.expiration}")
    private long expiration;

    public String generateToken(Long userId, String role) {
        return JWT.create()
                .withClaim("userId", userId)
                .withClaim("role", role)
                .withExpiresAt(new Date(System.currentTimeMillis() + expiration))
                .sign(Algorithm.HMAC256(secret));
    }

    public Long getUserId(String token) {
        return JWT.require(Algorithm.HMAC256(secret))
                .build().verify(token).getClaim("userId").asLong();
    }

    public String getRole(String token) {
        return JWT.require(Algorithm.HMAC256(secret))
                .build().verify(token).getClaim("role").asString();
    }
}
```

- [ ] **Step 3: 编写完整 init.sql**

```sql
-- 基于用户已有的表结构，补充 t_question_knowledge 和 t_mastery
-- 以及字段优化

CREATE TABLE IF NOT EXISTS t_question_knowledge (
    id BIGINT PRIMARY KEY AUTO_INCREMENT,
    question_id BIGINT NOT NULL,
    knowledge_id BIGINT NOT NULL,
    UNIQUE KEY uk_question_knowledge (question_id, knowledge_id)
);

CREATE TABLE IF NOT EXISTS t_mastery (
    id BIGINT PRIMARY KEY AUTO_INCREMENT,
    student_id BIGINT NOT NULL,
    knowledge_id BIGINT NOT NULL,
    mastery_score DECIMAL(5,2) DEFAULT 50.00,
    total_attempts INT DEFAULT 0,
    correct_count INT DEFAULT 0,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    UNIQUE KEY uk_student_knowledge (student_id, knowledge_id)
);

-- 给 t_question 补充字段 (if not exists)
ALTER TABLE t_question ADD COLUMN IF NOT EXISTS discrimination DECIMAL(3,2) DEFAULT 0.50 COMMENT '区分度';
ALTER TABLE t_question ADD COLUMN IF NOT EXISTS correct_rate DECIMAL(5,2) DEFAULT NULL COMMENT '历史正确率';

-- 给 t_exam 补充字段
ALTER TABLE t_exam ADD COLUMN IF NOT EXISTS strategy_json JSON COMMENT '组卷策略快照';
```

- [ ] **Step 4: 提交**

```bash
git add exam-system-backend/sql/ exam-system-backend/src/
git commit -m "feat: add JWT auth util and database migration scripts"
```

---

## Phase 3: 认证模块

### Task 3.1: 后端登录接口 + JWT 拦截器

**Files:**
- Create: `exam-system-backend/src/main/java/com/exam/module/auth/AuthController.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/user/entity/User.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/user/UserMapper.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/user/UserService.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/user/UserServiceImpl.java`
- Create: `exam-system-backend/src/main/java/com/exam/interceptor/JwtInterceptor.java`

- [ ] **Step 1: 创建 User 实体**

```java
package com.exam.module.user.entity;

import com.baomidou.mybatisplus.annotation.*;
import lombok.Data;
import java.time.LocalDateTime;

@Data
@TableName("t_user")
public class User {
    @TableId(type = IdType.AUTO)
    private Long id;
    private String username;
    private String password;
    private String realName;
    private String role; // ADMIN, TEACHER, STUDENT
    private String phone;
    private String email;
    @TableLogic
    private Integer deleted;
    @TableField(fill = FieldFill.INSERT)
    private LocalDateTime createTime;
    @TableField(fill = FieldFill.UPDATE)
    private LocalDateTime updateTime;
}
```

- [ ] **Step 2: 创建 UserMapper**

```java
package com.exam.module.user;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import com.exam.module.user.entity.User;

public interface UserMapper extends BaseMapper<User> {
}
```

- [ ] **Step 3: 创建登录接口**

```java
package com.exam.module.auth;

import lombok.Data;
import jakarta.validation.constraints.NotBlank;

@Data
public class LoginDTO {
    @NotBlank(message = "用户名不能为空")
    private String username;
    @NotBlank(message = "密码不能为空")
    private String password;
}
```

```java
@RestController
@RequestMapping("/api/auth")
public class AuthController {
    @Autowired
    private UserService userService;
    @Autowired
    private JwtUtil jwtUtil;

    @PostMapping("/login")
    public ApiResult<LoginVO> login(@Valid @RequestBody LoginDTO dto) {
        User user = userService.lambdaQuery()
                .eq(User::getUsername, dto.getUsername())
                .one();
        if (user == null || !user.getPassword().equals(dto.getPassword())) {
            return ApiResult.error(401, "用户名或密码错误");
        }
        String token = jwtUtil.generateToken(user.getId(), user.getRole());
        LoginVO vo = new LoginVO();
        vo.setToken(token);
        vo.setUserId(user.getId());
        vo.setUsername(user.getUsername());
        vo.setRealName(user.getRealName());
        vo.setRole(user.getRole());
        return ApiResult.success(vo);
    }
}
```

- [ ] **Step 4: 创建 JWT 拦截器**

```java
package com.exam.interceptor;

@Component
public class JwtInterceptor implements HandlerInterceptor {
    @Autowired
    private JwtUtil jwtUtil;

    @Override
    public boolean preHandle(HttpServletRequest request, HttpServletResponse response, Object handler) {
        String token = request.getHeader("Authorization");
        if (token == null || !token.startsWith("Bearer ")) {
            throw new BusinessException(401, "未登录");
        }
        token = token.substring(7);
        try {
            Long userId = jwtUtil.getUserId(token);
            String role = jwtUtil.getRole(token);
            request.setAttribute("userId", userId);
            request.setAttribute("role", role);
            return true;
        } catch (Exception e) {
            throw new BusinessException(401, "token无效或已过期");
        }
    }
}
```

- [ ] **Step 5: 注册拦截器 (WebMvcConfig)**

```java
package com.exam.common.config;

@Configuration
public class WebMvcConfig implements WebMvcConfigurer {
    @Autowired
    private JwtInterceptor jwtInterceptor;

    @Override
    public void addInterceptors(InterceptorRegistry registry) {
        registry.addInterceptor(jwtInterceptor)
                .addPathPatterns("/api/**")
                .excludePathPatterns("/api/auth/login");
    }
}
```

- [ ] **Step 6: 测试登录接口**

Run: `curl -X POST http://localhost:8080/api/auth/login -H "Content-Type: application/json" -d '{"username":"admin","password":"admin123"}'`
Expected: 返回 200 + token 数据

- [ ] **Step 7: 提交**

```bash
git add exam-system-backend/src/main/java/com/exam/
git commit -m "feat: implement auth module with JWT interceptor"
```

---

### Task 3.2: 前端登录页

**Files:**
- Create: `exam-system-frontend/src/views/login/LoginView.vue`
- Create: `exam-system-frontend/src/api/auth.js`
- Create: `exam-system-frontend/src/stores/user.js`
- Modify: `exam-system-frontend/src/router/index.js`

- [ ] **Step 1: 创建 auth API**

```javascript
// src/api/auth.js
import request from './request'

export function login(data) {
  return request.post('/auth/login', data)
}
```

- [ ] **Step 2: 创建 user store**

```javascript
// src/stores/user.js
import { defineStore } from 'pinia'
import { ref } from 'vue'

export const useUserStore = defineStore('user', () => {
  const token = ref(localStorage.getItem('token') || '')
  const userInfo = ref(JSON.parse(localStorage.getItem('userInfo') || '{}'))

  function setUserInfo(info) {
    userInfo.value = info
    localStorage.setItem('userInfo', JSON.stringify(info))
  }

  function setToken(t) {
    token.value = t
    localStorage.setItem('token', t)
  }

  function logout() {
    token.value = ''
    userInfo.value = {}
    localStorage.removeItem('token')
    localStorage.removeItem('userInfo')
  }

  return { token, userInfo, setUserInfo, setToken, logout }
})
```

- [ ] **Step 3: 创建登录页面**

```vue
<template>
  <div class="login-container">
    <div class="login-card">
      <h2 class="login-title">智能题库与自适应组卷系统</h2>
      <el-form :model="form" :rules="rules" ref="formRef" size="large">
        <el-form-item prop="username">
          <el-input v-model="form.username" placeholder="用户名" :prefix-icon="User" />
        </el-form-item>
        <el-form-item prop="password">
          <el-input v-model="form.password" type="password" placeholder="密码" :prefix-icon="Lock" show-password />
        </el-form-item>
        <el-form-item>
          <el-button type="primary" @click="handleLogin" :loading="loading" class="login-btn">登 录</el-button>
        </el-form-item>
      </el-form>
    </div>
  </div>
</template>

<script setup>
import { reactive, ref } from 'vue'
import { useRouter } from 'vue-router'
import { useUserStore } from '@/stores/user'
import { login } from '@/api/auth'
import { ElMessage } from 'element-plus'
import { User, Lock } from '@element-plus/icons-vue'

const router = useRouter()
const userStore = useUserStore()
const formRef = ref(null)
const loading = ref(false)

const form = reactive({
  username: '',
  password: ''
})

const rules = {
  username: [{ required: true, message: '请输入用户名', trigger: 'blur' }],
  password: [{ required: true, message: '请输入密码', trigger: 'blur' }]
}

async function handleLogin() {
  await formRef.value.validate()
  loading.value = true
  try {
    const res = await login(form)
    userStore.setToken(res.data.token)
    userStore.setUserInfo(res.data)
    ElMessage.success('登录成功')
    const roleMap = { ADMIN: '/admin', TEACHER: '/teacher', STUDENT: '/student' }
    router.push(roleMap[res.data.role] || '/')
  } finally {
    loading.value = false
  }
}
</script>

<style scoped>
.login-container {
  height: 100vh; display: flex; align-items: center; justify-content: center;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
}
.login-card {
  width: 420px; padding: 40px; background: #fff; border-radius: 12px; box-shadow: 0 20px 60px rgba(0,0,0,0.15);
}
.login-title { text-align: center; margin-bottom: 32px; color: #303133; font-size: 22px; font-weight: 600; }
.login-btn { width: 100%; }
</style>
```

- [ ] **Step 4: 配置路由守卫**

```javascript
// src/router/index.js (更新)
import { useUserStore } from '@/stores/user'

router.beforeEach((to, from, next) => {
  const userStore = useUserStore()
  if (to.path !== '/login' && !userStore.token) {
    next('/login')
  } else {
    next()
  }
})
```

- [ ] **Step 5: 提交**

```bash
git add exam-system-frontend/
git commit -m "feat: implement login page with auth flow"
```

---

## Phase 4: 用户管理 + 科目管理

### Task 4.1: 后端用户管理

**Files:**
- Create: `exam-system-backend/src/main/java/com/exam/module/user/UserController.java`

- [ ] **Step 1: 创建 UserController**

```java
@RestController
@RequestMapping("/api/user")
public class UserController {
    @Autowired
    private UserMapper userMapper;

    @GetMapping("/list")
    public ApiResult<IPage<User>> list(@RequestParam(defaultValue = "1") Integer page,
                                        @RequestParam(defaultValue = "10") Integer size,
                                        @RequestParam(required = false) String role,
                                        @RequestParam(required = false) String keyword) {
        Page<User> p = new Page<>(page, size);
        LambdaQueryWrapper<User> wrapper = new LambdaQueryWrapper<>();
        if (role != null) wrapper.eq(User::getRole, role);
        if (keyword != null) wrapper.like(User::getUsername, keyword).or().like(User::getRealName, keyword);
        return ApiResult.success(userMapper.selectPage(p, wrapper));
    }

    @PostMapping("/create")
    public ApiResult<Void> create(@Valid @RequestBody User user) {
        if (userMapper.selectCount(lambdaQuery().eq(User::getUsername, user.getUsername())) > 0) {
            return ApiResult.error(400, "用户名已存在");
        }
        userMapper.insert(user);
        return ApiResult.success(null);
    }

    @PutMapping("/update")
    public ApiResult<Void> update(@RequestBody User user) {
        userMapper.updateById(user);
        return ApiResult.success(null);
    }

    @DeleteMapping("/{id}")
    public ApiResult<Void> delete(@PathVariable Long id) {
        userMapper.deleteById(id);
        return ApiResult.success(null);
    }
}
```

- [ ] **Step 2: 提交**

---

### Task 4.2: 后端科目管理

**Files:**
- Create: `exam-system-backend/src/main/java/com/exam/module/subject/entity/Subject.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/subject/SubjectMapper.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/subject/SubjectController.java`

- [ ] **Step 1: 实现标准 CRUD (与用户管理类似)**

- [ ] **Step 2: 提交**

---

### Task 4.3: 前端管理员面板

**Files:**
- Create: `exam-system-frontend/src/layout/AdminLayout.vue`
- Create: `exam-system-frontend/src/views/admin/UserManage.vue`
- Create: `exam-system-frontend/src/views/admin/SubjectManage.vue`
- Modify: `exam-system-frontend/src/router/index.js`

- [ ] **Step 1: 创建管理员布局 (侧边栏 + 顶栏 + 内容区)**

```vue
<!-- AdminLayout.vue -->
<template>
  <el-container style="height: 100vh">
    <el-aside width="220px">
      <div class="logo">智能题库系统</div>
      <el-menu router :default-active="route.path" background-color="#304156" text-color="#bfcbd9" active-text-color="#409EFF">
        <el-menu-item index="/admin/users"><el-icon><User /></el-icon>用户管理</el-menu-item>
        <el-menu-item index="/admin/subjects"><el-icon><Collection /></el-icon>科目管理</el-menu-item>
        <el-menu-item index="/login" @click="userStore.logout()"><el-icon><SwitchButton /></el-icon>退出登录</el-menu-item>
      </el-menu>
    </el-aside>
    <el-container>
      <el-header style="background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,0.08); display: flex; align-items: center; justify-content: flex-end;">
        <span>{{ userStore.userInfo?.realName || userStore.userInfo?.username }}</span>
      </el-header>
      <el-main style="background: #f0f2f5">
        <router-view />
      </el-main>
    </el-container>
  </el-container>
</template>
```

- [ ] **Step 2: 创建用户管理页面 (el-table + 弹窗表单)**

```vue
<!-- UserManage.vue -->
<template>
  <div>
    <el-card>
      <div style="margin-bottom: 16px">
        <el-button type="primary" @click="openCreate">新增用户</el-button>
      </div>
      <el-table :data="users" border stripe v-loading="loading">
        <el-table-column prop="id" label="ID" width="80" />
        <el-table-column prop="username" label="用户名" />
        <el-table-column prop="realName" label="姓名" />
        <el-table-column prop="role" label="角色">
          <template #default="{ row }">{{ roleMap[row.role] }}</template>
        </el-table-column>
        <el-table-column prop="phone" label="手机号" />
        <el-table-column prop="email" label="邮箱" />
        <el-table-column label="操作" width="200">
          <template #default="{ row }">
            <el-button size="small" @click="openEdit(row)">编辑</el-button>
            <el-button size="small" type="danger" @click="handleDelete(row.id)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
      <el-pagination v-model:current-page="page" v-model:page-size="size" :total="total" layout="total, prev, pager, next" style="margin-top: 16px" />
    </el-card>

    <el-dialog v-model="dialogVisible" :title="isEdit ? '编辑用户' : '新增用户'" width="500px">
      <el-form :model="form" :rules="rules" ref="formRef" label-width="80px">
        <el-form-item label="用户名" prop="username"><el-input v-model="form.username" /></el-form-item>
        <el-form-item label="密码" prop="password"><el-input v-model="form.password" type="password" /></el-form-item>
        <el-form-item label="姓名" prop="realName"><el-input v-model="form.realName" /></el-form-item>
        <el-form-item label="角色" prop="role">
          <el-select v-model="form.role">
            <el-option label="管理员" value="ADMIN" />
            <el-option label="教师" value="TEACHER" />
            <el-option label="学生" value="STUDENT" />
          </el-select>
        </el-form-item>
        <el-form-item label="手机号"><el-input v-model="form.phone" /></el-form-item>
        <el-form-item label="邮箱"><el-input v-model="form.email" /></el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="dialogVisible = false">取消</el-button>
        <el-button type="primary" @click="handleSave">保存</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { getUsers, createUser, updateUser, deleteUser } from '@/api/user'
import { ElMessage, ElMessageBox } from 'element-plus'

const users = ref([])
const loading = ref(false)
const page = ref(1)
const size = ref(10)
const total = ref(0)
const dialogVisible = ref(false)
const isEdit = ref(false)
const formRef = ref(null)
const roleMap = { ADMIN: '管理员', TEACHER: '教师', STUDENT: '学生' }
const defaultForm = { username: '', password: '', realName: '', role: 'STUDENT', phone: '', email: '' }
const form = ref({ ...defaultForm })

async function fetchUsers() {
  loading.value = true
  try {
    const res = await getUsers({ page: page.value, size: size.value })
    users.value = res.data.records
    total.value = res.data.total
  } finally { loading.value = false }
}

function openCreate() {
  isEdit.value = false; form.value = { ...defaultForm }; dialogVisible.value = true
}

function openEdit(row) {
  isEdit.value = true; form.value = { ...row, password: '' }; dialogVisible.value = true
}

async function handleSave() {
  await formRef.value.validate()
  if (isEdit.value) await updateUser(form.value)
  else await createUser(form.value)
  ElMessage.success('操作成功')
  dialogVisible.value = false
  fetchUsers()
}

async function handleDelete(id) {
  await ElMessageBox.confirm('确认删除该用户？')
  await deleteUser(id)
  ElMessage.success('删除成功')
  fetchUsers()
}

const rules = {
  username: [{ required: true, message: '请输入用户名' }],
  password: [{ required: true, message: '请输入密码' }],
  realName: [{ required: true, message: '请输入姓名' }],
  role: [{ required: true, message: '请选择角色' }]
}

onMounted(fetchUsers)
</script>
```

- [ ] **Step 3: 创建科目管理页面 (类似结构，el-table + 弹窗)**

- [ ] **Step 4: 更新路由**

```javascript
const routes = [
  { path: '/login', component: () => import('@/views/login/LoginView.vue') },
  {
    path: '/admin',
    component: () => import('@/layout/AdminLayout.vue'),
    children: [
      { path: 'users', component: () => import('@/views/admin/UserManage.vue') },
      { path: 'subjects', component: () => import('@/views/admin/SubjectManage.vue') }
    ]
  }
]
```

- [ ] **Step 5: 提交**

---

## Phase 5: 题库管理

### Task 5.1: 后端题库 CRUD

**Files:**
- Create: `exam-system-backend/src/main/java/com/exam/module/question/entity/Question.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/question/entity/QuestionOption.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/question/QuestionMapper.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/question/QuestionController.java`

- [ ] **Step 1: 创建 Question 实体**

```java
@Data
@TableName("t_question")
public class Question {
    @TableId(type = IdType.AUTO)
    private Long id;
    private Long subjectId;
    private String type;       // SINGLE, MULTIPLE, FILL, SUBJECTIVE
    private String content;    // 题干（富文本HTML）
    private String options;    // JSON 字符串，存所有选项
    private String answer;     // 正确答案
    private Integer difficulty; // 1-5
    private BigDecimal discrimination; // 区分度
    private BigDecimal correctRate;   // 正确率
    private String analysis;  // 答案解析
    @TableLogic
    private Integer deleted;
    private LocalDateTime createTime;
    private LocalDateTime updateTime;
}
```

- [ ] **Step 2: 创建 QuestionController**

```java
@RestController
@RequestMapping("/api/question")
public class QuestionController {
    @Autowired
    private QuestionMapper questionMapper;

    @GetMapping("/page")
    public ApiResult<IPage<Question>> page(@RequestParam(defaultValue = "1") Integer page,
                                            @RequestParam(defaultValue = "10") Integer size,
                                            @RequestParam(required = false) Long subjectId,
                                            @RequestParam(required = false) String type,
                                            @RequestParam(required = false) Integer difficulty,
                                            @RequestParam(required = false) Long knowledgeId) {
        Page<Question> p = new Page<>(page, size);
        LambdaQueryWrapper<Question> wrapper = new LambdaQueryWrapper<>();
        if (subjectId != null) wrapper.eq(Question::getSubjectId, subjectId);
        if (type != null) wrapper.eq(Question::getType, type);
        if (difficulty != null) wrapper.eq(Question::getDifficulty, difficulty);
        // 如果按知识点筛选，关联 t_question_knowledge 表
        return ApiResult.success(questionMapper.selectPage(p, wrapper));
    }

    @PostMapping("/create")
    public ApiResult<Void> create(@RequestBody QuestionDTO dto) { /* 插入题目 + 选项 + 知识点关联 */ }

    @PutMapping("/update")
    public ApiResult<Void> update(@RequestBody QuestionDTO dto) { /* 更新题目 */ }

    @DeleteMapping("/{id}")
    public ApiResult<Void> delete(@PathVariable Long id) { /* 逻辑删除 */ }

    @GetMapping("/{id}")
    public ApiResult<Question> get(@PathVariable Long id) { /* 查询单题 */ }
}
```

- [ ] **Step 3: Excel 批量导入**

```java
@PostMapping("/import")
public ApiResult<Void> importExcel(@RequestParam("file") MultipartFile file) {
    Workbook workbook = new XSSFWorkbook(file.getInputStream());
    Sheet sheet = workbook.getSheetAt(0);
    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
        Row row = sheet.getRow(i);
        // 解析每一行：题型、题干、选项、答案、难度、知识点
    }
    return ApiResult.success(null);
}
```

- [ ] **Step 4: 提交**

---

### Task 5.2: 前端题库管理页面

**Files:**
- Create: `exam-system-frontend/src/views/teacher/QuestionManage.vue`
- Create: `exam-system-frontend/src/views/teacher/QuestionForm.vue`
- Create: `exam-system-frontend/src/api/question.js`

- [ ] **Step 1: 题目列表页（el-table 展示，支持筛选）**

包含：分页表格、顶部筛选栏（科目/题型/难度下拉）、新建/编辑/删除按钮

- [ ] **Step 2: 题目编辑页（区分题型表单）**

```
选择题：题干(富文本) + 选项动态列表 + 选择正确答案
填空题：题干(用____占位) + 答案
解答题：题干(富文本) + 参考答案(富文本)
```

- [ ] **Step 3: 提交**

---

## Phase 6: 知识点管理

### Task 6.1: 后端知识点模块

**Files:**
- Create: `exam-system-backend/src/main/java/com/exam/module/knowledge/entity/Knowledge.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/knowledge/KnowledgeMapper.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/knowledge/KnowledgeController.java`

- [ ] **Step 1: 树形结构查询**

```java
@RestController
@RequestMapping("/api/knowledge")
public class KnowledgeController {
    @GetMapping("/tree/{subjectId}")
    public ApiResult<List<KnowledgeVO>> tree(@PathVariable Long subjectId) {
        List<Knowledge> all = knowledgeMapper.selectList(lambdaQuery().eq(Knowledge::getSubjectId, subjectId));
        // 转为树形结构
        List<KnowledgeVO> tree = buildTree(all, 0L);
        return ApiResult.success(tree);
    }

    private List<KnowledgeVO> buildTree(List<Knowledge> all, Long parentId) {
        return all.stream()
                .filter(k -> k.getParentId().equals(parentId))
                .map(k -> {
                    KnowledgeVO vo = new KnowledgeVO(k);
                    vo.setChildren(buildTree(all, k.getId()));
                    return vo;
                }).collect(Collectors.toList());
    }
}
```

- [ ] **Step 2: 提交**

---

### Task 6.2: 前端知识点树页面

**Files:**
- Create: `exam-system-frontend/src/views/teacher/KnowledgeManage.vue`
- Create: `exam-system-frontend/src/api/knowledge.js`

- [ ] **Step 1: 使用 el-tree 展示知识树 + 右键增删改**

```vue
<template>
  <el-card>
    <el-button type="primary" @click="openCreate">新增根节点</el-button>
    <el-tree :data="treeData" :props="{ children: 'children', label: 'name' }" node-key="id" draggable @node-drop="handleDrop">
      <template #default="{ node, data }">
        <span>{{ data.name }}</span>
        <span style="float: right; margin-right: 20px">
          <el-button size="small" @click="openCreate(data.id)">新增子级</el-button>
          <el-button size="small" @click="openEdit(data)">编辑</el-button>
          <el-button size="small" type="danger" @click="handleDelete(data.id)">删除</el-button>
        </span>
      </template>
    </el-tree>
  </el-card>
</template>
```

- [ ] **Step 2: 提交**

---

## Phase 7: 智能组卷引擎（核心算法）

### Task 7.1: 后端组卷引擎

**Files:**
- Create: `exam-system-backend/src/main/java/com/exam/module/paper/engine/PaperGenerator.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/paper/PaperController.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/paper/entity/Paper.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/paper/PaperMapper.java`

- [ ] **Step 1: 实现三层组卷算法**

```java
@Component
public class PaperGenerator {

    @Autowired
    private QuestionMapper questionMapper;

    /**
     * 生成试卷
     * @param req 组卷请求（科目、题型数量、难度分布、知识点范围）
     * @param studentId 学生ID（用于自适应）
     */
    public List<Long> generate(PaperGenerateRequest req, Long studentId) {
        // 第一层：按科目+类型过滤
        List<Question> pool = questionMapper.selectList(buildBaseFilter(req));
        
        // 第二层：均衡分配（按知识点分组 -> 按难度分层 -> 按比例抽取）
        Map<Long, List<Question>> byKnowledge = pool.stream()
                .collect(Collectors.groupingBy(q -> getKnowledgeId(q.getId())));
        
        List<Long> selected = new ArrayList<>();
        for (Map.Entry<Long, List<Question>> entry : byKnowledge.entrySet()) {
            Map<Integer, List<Question>> byDifficulty = entry.getValue().stream()
                    .collect(Collectors.groupingBy(Question::getDifficulty));
            // 从各难度按比例抽题
            for (Map.Entry<Integer, List<Question>> diffEntry : byDifficulty.entrySet()) {
                int count = calculateCount(req, diffEntry.getKey(), entry.getKey());
                selected.addAll(randomPick(diffEntry.getValue(), count));
            }
        }

        // 第三层：能力适配（如果有学生ID）
        if (studentId != null) {
            selected = adaptToStudent(selected, studentId);
        }

        return selected;
    }

    // 能力适配：替换部分题目匹配学生水平
    private List<Long> adaptToStudent(List<Long> questionIds, Long studentId) {
        // 1. 查询学生各知识点掌握度
        // 2. 掌握度低的 -> 把该知识点的高难度题替换为低难度
        // 3. 保证替换比例不超过 40%
    }
}
```

- [ ] **Step 2: 组卷 API 接口**

```java
@PostMapping("/api/paper/generate")
public ApiResult<PaperVO> generate(@RequestBody PaperGenerateRequest req,
                                    @RequestAttribute Long userId,
                                    @RequestAttribute String role) {
    Long studentId = "STUDENT".equals(role) ? userId : null;
    List<Long> ids = paperGenerator.generate(req, studentId);
    // 创建 Paper 并保存
    return ApiResult.success(paperVO);
}
```

- [ ] **Step 3: 提交**

---

### Task 7.2: 前端组卷页面

**Files:**
- Create: `exam-system-frontend/src/views/teacher/PaperGenerate.vue`
- Create: `exam-system-frontend/src/views/teacher/PaperPreview.vue`

- [ ] **Step 1: 组卷配置表单**

```
选择科目 → 选择知识点（可多选）→ 设置各类题型数量
→ 设置难度分布（拖拽滑块调整比例：简单__% 中等__% 困难__%）
→ 点击"生成试卷"
```

- [ ] **Step 2: 试卷预览页**

展示完整的试卷内容，每题后面有"替换"按钮，可以手动换题

- [ ] **Step 3: 提交**

---

## Phase 8: 在线考试

### Task 8.1: 后端考试模块

**Files:**
- Create: `exam-system-backend/src/main/java/com/exam/module/exam/ExamController.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/exam/ExamService.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/exam/entity/Exam.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/exam/entity/AnswerRecord.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/exam/entity/AnswerDetail.java`

- [ ] **Step 1: 考试管理接口**

```
POST   /api/exam/create      — 创建考试（选择试卷+设置时间）
GET    /api/exam/list         — 考试列表
GET    /api/exam/student-list — 学生待考列表
```

- [ ] **Step 2: 答题接口**

```java
@PostMapping("/api/exam/submit")
public ApiResult<Void> submit(@RequestBody AnswerSubmitDTO dto, @RequestAttribute Long userId) {
    // 1. 创建 AnswerRecord
    // 2. 逐题判分（选择题自动判，解答题标记待批改）
    // 3. 更新学生掌握度（t_mastery）
    // 4. 更新题目正确率
}
```

- [ ] **Step 3: 提交**

---

### Task 8.2: 前端在线答题页

**Files:**
- Create: `exam-system-frontend/src/views/student/ExamList.vue`
- Create: `exam-system-frontend/src/views/student/ExamTaking.vue`

- [ ] **Step 1: 在线答题页（核心交互页面）**

```
顶部：考试标题 + 倒计时
左侧：答题导航（数字按钮，绿=已答，灰=未答）
右侧：题目区（显示题目内容 + 选项/输入框）
底部：上一题/下一题 + 提交按钮
```

- [ ] **Step 2: 提交

---

## Phase 9: 成绩管理与统计分析

### Task 9.1: 后端成绩与统计

**Files:**
- Create: `exam-system-backend/src/main/java/com/exam/module/score/ScoreController.java`
- Create: `exam-system-backend/src/main/java/com/exam/module/statistics/StatisticsController.java`

- [ ] **Step 1: 成绩接口**

```
GET /api/score/list?examId=     — 某次考试的成绩列表
GET /api/score/detail?recordId= — 答卷详情（每题得分）
GET /api/score/export?examId=   — 导出 Excel
```

- [ ] **Step 2: 统计接口**

```java
@GetMapping("/api/statistics/class-overview")
public ApiResult<ClassOverviewVO> classOverview(@RequestParam Long examId) {
    // 平均分 / 及格率 / 优秀率 / 最高分 / 最低分
}

@GetMapping("/api/statistics/knowledge-mastery")
public ApiResult<List<KnowledgeMasteryVO>> knowledgeMastery(@RequestParam Long examId) {
    // 各知识点全班平均掌握率 → 雷达图数据
}

@GetMapping("/api/statistics/student-trend")
public ApiResult<List<TrendVO>> studentTrend(@RequestParam Long studentId, @RequestParam Long subjectId) {
    // 学生历次成绩折线图数据
}
```

- [ ] **Step 3: 提交**

---

### Task 9.2: 前端图表展示

**Files:**
- Create: `exam-system-frontend/src/views/teacher/ScoreManage.vue`
- Create: `exam-system-frontend/src/views/teacher/StatisticsView.vue`
- Create: `exam-system-frontend/src/views/student/StudentScore.vue`
- Create: `exam-system-frontend/src/views/student/StudentAnalysis.vue`

- [ ] **Step 1: 学情分析页（三张 ECharts 图表）**

```vue
<template>
  <el-row :gutter="20">
    <el-col :span="12">
      <el-card>
        <h3>成绩分布</h3>
        <v-chart :option="distributionOption" style="height: 300px" />
      </el-card>
    </el-col>
    <el-col :span="12">
      <el-card>
        <h3>知识点掌握雷达图</h3>
        <v-chart :option="radarOption" style="height: 300px" />
      </el-card>
    </el-col>
  </el-row>
  <el-row style="margin-top: 20px">
    <el-col :span="24">
      <el-card>
        <h3>薄弱知识点 TOP5</h3>
        <el-table :data="weakKnowledge" stripe>
          <el-table-column prop="name" label="知识点" />
          <el-table-column prop="correctRate" label="正确率">
            <template #default="{ row }">
              <el-progress :percentage="row.correctRate" :status="row.correctRate < 60 ? 'exception' : 'success'" />
            </template>
          </el-table-column>
        </el-table>
      </el-card>
    </el-col>
  </el-row>
</template>
```

- [ ] **Step 2: 提交**

---

## Phase 10: 自适应测评

### Task 10.1: 后端自适应逻辑

**Files:**
- Create: `exam-system-backend/src/main/java/com/exam/module/paper/MasteryService.java`
- Create: `exam-system-frontend/src/views/student/WrongBook.vue`
- Create: `exam-system-frontend/src/views/student/AdaptivePractice.vue`

- [ ] **Step 1: 掌握度更新服务**

```java
@Service
public class MasteryService {
    @Autowired
    private MasteryMapper masteryMapper;

    public void updateMastery(Long studentId, Long knowledgeId, boolean isCorrect) {
        Mastery mastery = masteryMapper.selectOne(lambdaQuery()
                .eq(Mastery::getStudentId, studentId)
                .eq(Mastery::getKnowledgeId, knowledgeId));
        if (mastery == null) {
            mastery = new Mastery();
            mastery.setStudentId(studentId);
            mastery.setKnowledgeId(knowledgeId);
            mastery.setMasteryScore(BigDecimal.valueOf(50));
        }
        // 递推公式：new = old × 0.7 + correct × 0.3
        int correctScore = isCorrect ? 100 : 0;
        BigDecimal newScore = mastery.getMasteryScore()
                .multiply(BigDecimal.valueOf(0.7))
                .add(BigDecimal.valueOf(correctScore).multiply(BigDecimal.valueOf(0.3)));
        mastery.setMasteryScore(newScore);
        mastery.setTotalAttempts(mastery.getTotalAttempts() + 1);
        if (isCorrect) mastery.setCorrectCount(mastery.getCorrectCount() + 1);
        masteryMapper.insertOrUpdate(mastery);
    }

    public List<KnowledgeMasteryVO> getStudentMastery(Long studentId) {
        // 查询学生对所有知识点的掌握度
    }
}
```

- [ ] **Step 2: 错题本+针对性推荐**

```java
@GetMapping("/api/question/wrong-book")
public ApiResult<List<Question>> wrongBook(@RequestParam Long studentId) {
    // 查询该生答错的题目，按知识点分组
}

@GetMapping("/api/question/recommend")
public ApiResult<List<Question>> recommend(@RequestParam Long studentId) {
    // 根据薄弱知识点推荐题目
}
```

- [ ] **Step 3: 提交**

---

## 未实现功能（后续升级方向）

- IRT 项目反应理论引擎 (`IRTEngine.java`)
- 考试中动态自适应出题（而非考后分析）
- 移动端 App（基于同一套 API）
- 系统操作日志
- 消息通知（考试提醒等）

---

## 实现顺序总结

| 阶段 | 内容 | 交付物 |
|------|------|--------|
| Phase 1 | 前后端脚手架 | 可启动的空项目 |
| Phase 2 | 数据库 + JWT | 基础认证可用 |
| Phase 3 | 登录 | 可登录/登出/路由守卫 |
| Phase 4 | 用户+科目管理 | 管理员端可用 |
| Phase 5 | 题库管理 | 教师可管理题目 |
| Phase 6 | 知识点管理 | 树形知识点维护 |
| Phase 7 | 智能组卷 | 核心算法可用 |
| Phase 8 | 在线考试 | 学生可考试 |
| Phase 9 | 成绩+统计 | 图表展示 |
| Phase 10 | 自适应 | 错题本+推荐 |
