"""
配置加载模块
加载供应商和产品映射配置
"""

import yaml
from typing import Dict, Any, Optional


class ConfigLoader:
    """配置加载器"""

    def __init__(self, config_path: str):
        self.config_path = config_path
        self.config = {}
        self._load_config()

    def _load_config(self):
        """加载YAML配置文件"""
        with open(self.config_path, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)

    def get_supplier_info(self, supplier_code: str) -> Optional[Dict[str, Any]]:
        """获取供应商信息"""
        suppliers = self.config.get("suppliers", {})
        return suppliers.get(supplier_code)

    def get_product_info(self, product_code: str) -> Optional[Dict[str, Any]]:
        """获取产品信息"""
        products = self.config.get("product_categories", {})
        return products.get(product_code)

    def get_all_suppliers(self) -> Dict[str, Any]:
        """获取所有供应商"""
        return self.config.get("suppliers", {})

    def get_all_products(self) -> Dict[str, Any]:
        """获取所有产品"""
        return self.config.get("product_categories", {})

    def get_template_path(self, template_type: str) -> str:
        """获取模板路径"""
        templates = {
            "contract": self.config.get("contract_templates", {}).get("default", ""),
            "customs_declaration": self.config.get("customs_declaration_templates", {}).get("default", ""),
            "declaration_elements": self.config.get("declaration_element_templates", {}).get("default", ""),
        }
        return templates.get(template_type, "")

    def get_export_declaration_template_filename(self) -> str:
        """
        出口申报要素 Word 壳文件名（位于 export_templates/，与进口 templates/ 下的申报要素模板不同）。
        """
        return self.config.get("export_declaration_element_templates", {}).get(
            "default", "申报要素总汇.docx"
        )

    def match_supplier_by_name(self, supplier_name: str) -> Optional[tuple]:
        """
        根据供应商名称匹配供应商代码
        返回 (supplier_code, supplier_info) 或 None
        """
        suppliers = self.config.get("suppliers", {})
        supplier_name_lower = supplier_name.lower()
        for code, info in suppliers.items():
            supplier_name_cfg = info.get("name", "").lower()
            # 匹配全称或简称
            if supplier_name_cfg in supplier_name_lower or supplier_name_lower in supplier_name_cfg:
                return code, info
            # 特殊处理Alcantara (包括Code: ALCANTARA SPA格式)
            if "alcantara" in supplier_name_lower and "alcantara" in supplier_name_cfg:
                return code, info
            if "alcantara spa" in supplier_name_lower:
                return "alcantara", suppliers.get("alcantara")
            # 特殊处理Crest Leather
            if "crest" in supplier_name_lower and "crest" in supplier_name_cfg:
                return code, info
            # 匹配Crest JMT Leather Ltd
            if "jmt" in supplier_name_lower and "jmt" in supplier_name_cfg:
                return code, info
            # 特殊处理DECA GLOBAL
            if "deca" in supplier_name_lower and "deca" in supplier_name_cfg:
                return code, info
            # 特殊处理WIPELLI
            if "wipelli" in supplier_name_lower and "wipelli" in supplier_name_cfg:
                return code, info
            # 特殊处理Mastrotto
            if "mastrotto" in supplier_name_lower and "mastrotto" in supplier_name_cfg:
                return code, info
            # 特殊处理West Trading
            if "west trading" in supplier_name_lower and "west trading" in supplier_name_cfg:
                return code, info
            # 特殊处理Konrad Hornschuch
            if "hornschuch" in supplier_name_lower and "hornschuch" in supplier_name_cfg:
                return code, info
            if "konrad" in supplier_name_lower and "konrad" in supplier_name_cfg:
                return code, info
            if "continental" in supplier_name_lower:
                return "konrad_hornschuch", suppliers.get("konrad_hornschuch", {})
            # 特殊处理MABO
            if "mabo" in supplier_name_lower and "mabo" in supplier_name_cfg:
                return code, info
            # 特殊处理HDM
            if "higher dimension" in supplier_name_lower and "higher dimension" in supplier_name_cfg:
                return code, info
            if "hdm" in supplier_name_lower and "hdm" in supplier_name_cfg:
                return code, info
        return None

    def match_product_by_name(self, product_name: str) -> Optional[tuple]:
        """
        根据产品名称匹配产品代码
        返回 (product_code, product_info) 或 None
        """
        products = self.config.get("product_categories", {})
        for code, info in products.items():
            if info.get("name", "") in product_name or product_name in info.get("name", ""):
                return code, info
        return None

    def match_product_by_keywords(self, description: str, supplier_code: str = None) -> Optional[tuple]:
        """
        根据商品描述关键词匹配产品，可选根据供应商过滤
        """
        products = self.config.get("product_categories", {})
        description_lower = description.lower()

        for code, info in products.items():
            product_name = info.get("name", "").lower()
            category = info.get("category", "").lower()
            sub_category = info.get("sub_category", "").lower()
            product_supplier = info.get("supplier", "")

            # 如果指定了供应商，先检查供应商匹配
            if supplier_code and product_supplier:
                # 支持supplier为列表或字符串
                if isinstance(product_supplier, list):
                    if supplier_code not in product_supplier:
                        continue
                elif supplier_code != product_supplier:
                    continue

            # 检查关键词匹配
            keywords = [product_name, category, sub_category]
            for keyword in keywords:
                if keyword and keyword in description_lower:
                    return code, info

            # 特殊处理Alcantara产品
            if "alcantara" in description_lower or "®" in description:
                if "alcantara" in code or "alcantara" in product_name:
                    return code, info

            # 特殊处理牛皮产品
            if "leather" in description_lower or "牛皮" in description or "hide" in description_lower:
                if "leather" in code or "牛皮" in product_name:
                    return code, info

            # 特殊处理MABO供应商 - 根据描述决定产品类型
            if supplier_code == "mabo":
                if "alcantara" in description_lower or "microfibre" in description_lower:
                    return "mabo_alcantara_fabric", products.get("mabo_alcantara_fabric", {})
                if "leather" in description_lower or "牛皮" in description_lower:
                    return "mabo_cowhide", products.get("mabo_cowhide", {})

            # 特殊处理HDM供应商 - 涂层织物
            if supplier_code == "hdm":
                if "coating" in description_lower or "涂层" in description_lower or "hex" in description_lower:
                    return "coated_fabric_hdm", products.get("coated_fabric_hdm", {})

        return None

    def match_product_by_item_code(self, item_code: str) -> Optional[tuple]:
        """
        根据产品代码前缀匹配产品
        Alcantara产品根据item code前4位决定申报要素
        """
        products = self.config.get("product_categories", {})

        # 提取item code前4位
        item_prefix = item_code[:4] if len(item_code) >= 4 else item_code

        for code, info in products.items():
            # 检查是否有item_code_prefix配置
            prefix = info.get("item_code_prefix")
            if prefix:
                if isinstance(prefix, list):
                    if item_prefix in prefix:
                        return code, info
                elif item_prefix == prefix:
                    return code, info

        return None

    def get_hs_code_by_item_code(self, item_code: str, supplier_code: str = None) -> str:
        """
        根据item code获取HS code
        Alcantara规则: 前4位为5012或5015/5030/5010时，HS为5603149000
        """
        products = self.config.get("product_categories", {})

        # 先尝试按item code前缀匹配
        match = self.match_product_by_item_code(item_code)
        if match:
            _, info = match
            return info.get("hs_code", "")

        # 如果没匹配到，返回空
        return ""

    def get_declaration_elements(self, product_code: str) -> Dict[str, Any]:
        """获取产品的申报要素"""
        products = self.config.get("product_categories", {})
        product_info = products.get(product_code, {})
        return product_info.get("declaration_elements", {})

    def match_fabric_by_composition(self, description: str, supplier_code: str = None) -> Optional[tuple]:
        """
        根据成分描述匹配面料产品
        West Trading规则:
        - 62% wool / 38% pes → HS 5112200000
        - 92% wool / 8% pes → HS 5112190000
        - 100% pes → HS 5407520091
        """
        products = self.config.get("product_categories", {})
        description_lower = description.lower()

        # 按优先级匹配（从最具体到最通用）
        composition_map = [
            ("62% wool / 38% pes", "wool_fabric_62_38"),
            ("62% wool / 38% pes", "wool_fabric_62_38"),
            ("92% wool / 8% pes", "wool_fabric_92_8"),
            ("100% pes", "polyester_fabric_100"),
            ("100% polyester", "polyester_fabric_100"),
        ]

        for pattern, product_code in composition_map:
            if pattern.lower() in description_lower:
                product_info = products.get(product_code, {})
                if product_info:
                    # 如果指定了供应商，检查供应商匹配
                    if supplier_code:
                        product_supplier = product_info.get("supplier", "")
                        if isinstance(product_supplier, list):
                            if supplier_code not in product_supplier:
                                continue
                        elif supplier_code != product_supplier:
                            continue
                    return product_code, product_info

        return None

    def get_fabric_hs_code_by_composition(self, description: str) -> str:
        """
        根据成分描述获取HS Code
        """
        match = self.match_fabric_by_composition(description)
        if match:
            _, info = match
            return info.get("hs_code", "")
        return ""
